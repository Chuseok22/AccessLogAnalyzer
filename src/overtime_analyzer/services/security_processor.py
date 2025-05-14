"""
경비 기록 처리 서비스

경비 기록 엑셀 파일을 처리하여 날짜별 경비 상태를 분석하는 기능을 제공합니다.
"""

import pandas as pd
import traceback
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime, time, date, timedelta

from ..models.data_models import SecurityStatusChange, SecurityPeriod, SecurityRecord
from ..utils.date_utils import calculate_business_date


class SecurityProcessor:
    """경비 기록 처리 클래스"""

    def __init__(self):
        self.unclear_security_days = []  # 경비 기록이 불명확한 날짜 목록

    def _parse_security_record(self, row: pd.Series, col_mapping: Dict[str, str]) -> SecurityRecord:
        """
        데이터프레임의 행을 SecurityRecord 객체로 변환합니다.

        Args:
            row: 변환할 데이터프레임 행
            col_mapping: 컬럼 매핑 정보

        Returns:
            SecurityRecord: 경비 기록 객체
        """
        record_date = row[col_mapping["발생일자"]].date()
        record_time = time(row["시간_시"], row["시간_분"])
        mode = row[col_mapping["모드"]]
        business_date = calculate_business_date(record_date, record_time)
        record_type = self._determine_record_type(row, col_mapping)

        return SecurityRecord(
            record_date=record_date,
            record_time=record_time,
            mode=mode,
            business_date=business_date,
            record_type=record_type,
        )

    def process_security_log(
        self, df: pd.DataFrame, start_date: str, end_date: str
    ) -> Dict[date, List[Dict[str, Any]]]:
        """
        경비 기록을 처리하여 각 날짜별 경비 상태를 분석합니다.

        Args:
            df: 경비 기록 데이터프레임
            start_date: 시작 날짜 (YYYY-MM-DD 형식)
            end_date: 종료 날짜 (YYYY-MM-DD 형식)

        Returns:
            Dict[date, List[Dict]]: 날짜별 경비 상태 변화 목록

        Raises:
            ValueError: 필수 컬럼을 찾을 수 없는 경우
        """
        try:
            # 필수 컬럼 정의
            required_cols = ["발생일자", "발생시각", "모드"]

            # 열 이름으로 컬럼 찾기
            col_mapping = self._find_columns(df)

            # 필요한 컬럼이 없으면 오류 반환
            missing_cols = []
            if col_mapping["모드"] is None:
                missing_cols.append("모드")
            if missing_cols:
                raise ValueError(f"다음 컬럼을 찾을 수 없습니다: {', '.join(missing_cols)}")

            # 날짜/시간 데이터 처리
            df = self._process_datetime_columns(df, col_mapping)

            # 필터링 적용
            filtered_df = self._apply_date_filter(df, col_mapping, start_date, end_date)

            # 필요한 열만 선택하여 메모리 사용 최적화
            filtered_df_slim = filtered_df[
                [
                    col_mapping["발생일자"],
                    col_mapping["발생시각"],
                    col_mapping["모드"],
                    "시간_시",
                    "시간_분",
                ]
            ].copy()

            # 이전날짜와 다음날짜 관계 분석을 위해 날짜별로 정렬
            filtered_df_slim = filtered_df_slim.sort_values(
                by=[col_mapping["발생일자"], col_mapping["발생시각"]]
            )

            # 각 기록에 시간대 태그 생성 (새벽 4시 기준)
            filtered_df_slim["시간대"] = "주간"
            filtered_df_slim.loc[
                (filtered_df_slim["시간_시"] >= 0) & (filtered_df_slim["시간_시"] < 4), "시간대"
            ] = "새벽"

            # 업무일 계산 (새벽 4시 기준)
            filtered_df_slim["업무일"] = filtered_df_slim[col_mapping["발생일자"]].dt.date
            # 새벽 시간대(0-4시)는 전날의 업무일로 계산
            filtered_df_slim.loc[filtered_df_slim["시간대"] == "새벽", "업무일"] = (
                filtered_df_slim.loc[
                    filtered_df_slim["시간대"] == "새벽", col_mapping["발생일자"]
                ].dt.date
                - pd.Timedelta(days=1)
            )

            # 각 기록 유형 판단
            filtered_df_slim["기록유형"] = filtered_df_slim.apply(
                lambda row: self._determine_record_type(row, col_mapping), axis=1
            )

            # 출입(불명확) 기록 처리 및 의심 데이터 감지
            business_days = filtered_df_slim["업무일"].unique()
            self._process_unclear_records(filtered_df_slim, business_days, col_mapping)

            # 각 업무일별 경비 상태 시간 분석
            security_status_by_day = self._analyze_security_status_by_day(
                filtered_df_slim, business_days, col_mapping
            )

            return security_status_by_day

        except Exception as e:
            # 예외 발생 시 상세 정보 출력하고 다시 발생
            print(f"경비 기록 처리 중 오류: {str(e)}")
            print(traceback.format_exc())
            raise

    def _find_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        데이터프레임에서 필요한 컬럼을 찾습니다.

        Args:
            df: 분석할 데이터프레임

        Returns:
            Dict[str, str]: 컬럼 매핑 정보
        """
        col_mapping = {}

        # 컬럼 이름 매핑 (정확한 이름 또는 포함된 문자열로 찾기)
        for col in df.columns:
            col_str = str(col).lower()  # 컬럼명을 소문자로 변환하여 비교
            if "발생일자" in col_str or "날짜" in col_str:
                col_mapping["발생일자"] = col
            elif "발생시각" in col_str or "시간" in col_str:
                col_mapping["발생시각"] = col
            elif "모드" in col_str or "상태" in col_str or "내용" in col_str:
                col_mapping["모드"] = col

        # 찾지 못한 컬럼은 기본 위치로 설정
        if "발생일자" not in col_mapping:
            col_mapping["발생일자"] = df.columns[0]  # A열
        if "발생시각" not in col_mapping:
            col_mapping["발생시각"] = df.columns[1]  # B열
        if "모드" not in col_mapping:
            if len(df.columns) > 8:
                col_mapping["모드"] = df.columns[8]  # I열
            else:
                col_mapping["모드"] = None

        return col_mapping

    def _process_datetime_columns(
        self, df: pd.DataFrame, col_mapping: Dict[str, str]
    ) -> pd.DataFrame:
        """
        날짜와 시간 열을 적절한 형식으로 변환합니다.

        Args:
            df: 처리할 데이터프레임
            col_mapping: 컬럼 매핑 정보

        Returns:
            pd.DataFrame: 처리된 데이터프레임
        """
        # 데이터프레임에서 날짜 열이 문자열이면 datetime으로 변환
        if not pd.api.types.is_datetime64_any_dtype(df[col_mapping["발생일자"]]):
            df[col_mapping["발생일자"]] = pd.to_datetime(
                df[col_mapping["발생일자"]], errors="coerce"
            )

        # 시각 열이 문자열이면 datetime.time으로 변환
        if not pd.api.types.is_datetime64_any_dtype(df[col_mapping["발생시각"]]):
            # 시각 데이터를 시간 및 분으로 분리하여 저장 (이후 분석용)
            df["시각_임시"] = pd.to_datetime(
                df[col_mapping["발생시각"]], errors="coerce", format="%H:%M:%S"
            )
            df["시간_시"] = df["시각_임시"].dt.hour
            df["시간_분"] = df["시각_임시"].dt.minute
            df = df.drop(columns=["시각_임시"])
        else:
            # 이미 datetime 형식인 경우 직접 시/분 추출
            df["시간_시"] = df[col_mapping["발생시각"]].dt.hour
            df["시간_분"] = df[col_mapping["발생시각"]].dt.minute

        return df

    def _determine_business_date(self, dt: datetime) -> date:
        """
        날짜와 시간 정보를 바탕으로 업무일을 결정합니다.
        새벽 4시 이전은 전날의 업무일로 처리합니다.

        Args:
            dt: 날짜 및 시간 정보

        Returns:
            date: 업무일
        """
        business_date = dt.date()
        if dt.hour < 4:  # 새벽 4시 이전은 전날 업무일로 처리
            business_date = business_date - timedelta(days=1)
        return business_date

    def _apply_date_filter(
        self, df: pd.DataFrame, col_mapping: Dict[str, str], start_date: str, end_date: str
    ) -> pd.DataFrame:
        """
        날짜 필터를 적용합니다.

        Args:
            df: 처리할 데이터프레임
            col_mapping: 컬럼 매핑 정보
            start_date: 시작 날짜 (YYYY-MM-DD 형식)
            end_date: 종료 날짜 (YYYY-MM-DD 형식)

        Returns:
            pd.DataFrame: 필터링된 데이터프레임
        """
        if start_date and end_date:
            filtered_df = df[
                (
                    df[col_mapping["발생일자"]].dt.date
                    >= datetime.strptime(start_date, "%Y-%m-%d").date()
                )
                & (
                    df[col_mapping["발생일자"]].dt.date
                    <= datetime.strptime(end_date, "%Y-%m-%d").date()
                )
            ]
            return filtered_df
        return df

    def _determine_record_type(self, row: pd.Series, col_mapping: Dict[str, str]) -> str:
        """
        기록을 경비해제/경비시작으로 판단합니다.

        Args:
            row: 판단할 데이터 행
            col_mapping: 컬럼 매핑 정보

        Returns:
            str: 기록 유형 ("경비해제", "경비시작", "출입(불명확)" 또는 "기타")
        """
        mode = str(row[col_mapping["모드"]]).lower()

        # 명시적인 경비 해제/설정 상태 확인
        if "출근" in mode or "해제" in mode:
            return "경비해제"
        elif "퇴근" in mode or "세팅" in mode or "세트" in mode:
            return "경비시작"
        elif "출입" in mode:
            # 출입은 맥락에 따라 나중에 해제/시작으로 판단
            return "출입(불명확)"

        return "기타"

    def _process_unclear_records(
        self, df: pd.DataFrame, business_days: List[date], col_mapping: Dict[str, str]
    ) -> None:
        """
        출입(불명확) 기록을 처리합니다.

        Args:
            df: 처리할 데이터프레임
            business_days: 업무일 목록
            col_mapping: 컬럼 매핑 정보
        """
        # 경비 데이터가 불명확한 업무일 기록
        self.unclear_security_days = []

        for business_day in business_days:
            day_records = df[df["업무일"] == business_day]
            unclear_records = day_records[day_records["기록유형"] == "출입(불명확)"]

            # 업무일 내 기록 시간순 정렬
            day_records_sorted = day_records.sort_values(
                by=[col_mapping["발생일자"], col_mapping["발생시각"]]
            )

            # 경비 해제/시작 기록 확인
            has_release = any(
                record["기록유형"] == "경비해제" for _, record in day_records_sorted.iterrows()
            )
            has_start = any(
                record["기록유형"] == "경비시작" for _, record in day_records_sorted.iterrows()
            )

            if len(unclear_records) > 0:
                # 불명확한 기록이 있고 명확한 기록이 없는 경우 의심 데이터로 분류
                if not has_release and not has_start:
                    self.unclear_security_days.append(
                        {
                            "날짜": business_day,
                            "기록수": len(day_records),
                            "불명확기록수": len(unclear_records),
                            "설명": "명확한 경비 해제/시작 기록 없음",
                        }
                    )

            # 명확한 경비 기록이 아예 없는 경우도 의심 데이터로 분류
            if not has_release and not has_start and len(day_records) > 0:
                if business_day not in [d["날짜"] for d in self.unclear_security_days]:
                    self.unclear_security_days.append(
                        {
                            "날짜": business_day,
                            "기록수": len(day_records),
                            "불명확기록수": 0,
                            "설명": "경비 해제/시작 기록 없음",
                        }
                    )

    def _analyze_security_status_by_day(
        self, df: pd.DataFrame, business_days: List[date], col_mapping: Dict[str, str]
    ) -> Dict[date, List[Dict[str, str]]]:
        """
        각 업무일별 경비 상태 시간을 분석합니다.

        Args:
            df: 처리할 데이터프레임
            business_days: 업무일 목록
            col_mapping: 컬럼 매핑 정보

        Returns:
            Dict[date, List[Dict]]: 날짜별 경비 상태 변화 목록
        """
        security_status_by_day = {}

        for business_day in business_days:
            day_records = df[df["업무일"] == business_day]
            day_records_sorted = day_records.sort_values(
                by=[col_mapping["발생일자"], col_mapping["발생시각"]]
            )

            # 해당 업무일의 경비 상태 시간 기록 초기화
            security_status = []

            for _, record in day_records_sorted.iterrows():
                # 시간 문자열 생성 (YYYY-MM-DD HH:MM 형식)
                record_datetime = record[col_mapping["발생일자"]].strftime("%Y-%m-%d")
                record_time = f"{record['시간_시']:02d}:{record['시간_분']:02d}"
                record_datetime_full = f"{record_datetime} {record_time}"
                record_type = record["기록유형"]

                # 경비해제/경비시작 시간 기록
                if record_type == "경비해제":
                    security_status.append({"시간": record_datetime_full, "상태": "해제"})
                elif record_type == "경비시작":
                    security_status.append({"시간": record_datetime_full, "상태": "시작"})

            # 업무일별 경비 상태 저장
            security_status_by_day[business_day] = security_status

        return security_status_by_day
