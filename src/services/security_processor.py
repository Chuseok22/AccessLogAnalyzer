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

    def process_security_log(
        self,
        df: pd.DataFrame,
        start_date: str = None,
        end_date: str = None,
    ) -> Dict[date, Dict[str, Any]]:
        """
        경비 기록 엑셀 파일을 처리하여 날짜별 경비 상태를 분석합니다.

        Args:
            df (pd.DataFrame): 경비 기록 데이터프레임
            start_date (str): 분석 시작 날짜 (YYYY-MM-DD)
            end_date (str): 분석 종료 날짜 (YYYY-MM-DD)

        Returns:
            Dict: 날짜별 경비 상태 및 기간 정보를 포함하는 사전
        """
        try:
            print("[DEBUG] 경비 기록 처리 시작")

            # 결과 사전 초기화
            security_data = {}

            # 데이터프레임이 비어있는 경우
            if df.empty:
                print("[WARNING] 경비 기록 데이터가 비어있습니다.")
                return security_data

            # 필요한 열이 있는지 확인
            required_columns = ["발생일자", "발생시각", "모드", "sensor_id", "출입자"]
            missing_columns = [col for col in required_columns if col not in df.columns]

            if missing_columns:
                print(f"[WARNING] 필요한 열이 없습니다: {missing_columns}")

                # 열 이름 목록 출력
                print(f"[DEBUG] 가능한 열 이름: {df.columns.tolist()}")

                # 비슷한 열 이름 찾기
                column_mapping = {
                    "발생일자": ["날짜", "일자", "일시", "발생일", "dttm", "date"],
                    "발생시각": ["시각", "시간", "time", "발생시간"],
                    "모드": ["상태", "mode", "type", "유형", "구분"],
                    "sensor_id": ["센서", "sensor", "id", "센서ID", "센서번호"],
                    "출입자": ["사용자", "user", "이름", "name", "person"],
                }

                # 열 이름 매핑 시도
                for req_col, alternatives in column_mapping.items():
                    if req_col in missing_columns:
                        for alt in alternatives:
                            if alt in df.columns:
                                df[req_col] = df[alt]
                                print(f"[INFO] '{alt}' 열을 '{req_col}'로 매핑했습니다.")
                                missing_columns.remove(req_col)
                                break

            # 여전히 필요한 열이 없는 경우
            if missing_columns:
                print(f"[ERROR] 필수 열이 없어 처리할 수 없습니다: {missing_columns}")
                return security_data

            # 날짜 및 시간 열 처리
            df = self._preprocess_datetime_columns(df)

            # 레코드 변환
            records = self._convert_to_security_records(df)

            # 날짜 필터링
            if start_date and end_date:
                start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
                end_date = datetime.strptime(end_date, "%Y-%m-%d").date()
                records = [r for r in records if start_date <= r.business_date <= end_date]

            # 날짜별로 레코드 그룹화
            date_records = {}
            for record in records:
                if record.business_date not in date_records:
                    date_records[record.business_date] = []
                date_records[record.business_date].append(record)

            # 각 날짜별 경비 상태 변화 분석
            for biz_date, day_records in date_records.items():
                # 시간순으로 정렬
                day_records.sort(key=lambda x: x.record_time)

                # 경비 상태 변화 추적
                status_changes = self._track_security_status_changes(day_records)

                # 경비 해제/설정 기간 도출
                security_periods = self._derive_security_periods(status_changes)

                # 결과 저장
                security_data[biz_date] = {
                    "records": day_records,
                    "status_changes": status_changes,
                    "periods": security_periods,
                }

                # 경비 상태가 불명확한 날짜 체크
                if not security_periods:
                    self.unclear_security_days.append(biz_date)

            print(f"[DEBUG] 경비 기록 처리 완료: {len(security_data)} 일자 처리됨")
            if self.unclear_security_days:
                print(f"[WARNING] {len(self.unclear_security_days)}일의 경비 상태가 불명확합니다.")

            return security_data

        except Exception as e:
            print(f"[ERROR] 경비 기록 처리 중 오류 발생: {str(e)}")
            traceback.print_exc()
            return {}

    def _preprocess_datetime_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        날짜 및 시간 열을 전처리합니다.

        Args:
            df (pd.DataFrame): 처리할 데이터프레임

        Returns:
            pd.DataFrame: 전처리된 데이터프레임
        """
        # 복사본 생성
        df = df.copy()

        try:
            # 발생일자 열이 문자열인 경우 처리
            if df["발생일자"].dtype == object:
                df["발생일자_parsed"] = pd.to_datetime(df["발생일자"], errors="coerce")
            else:
                df["발생일자_parsed"] = pd.to_datetime(df["발생일자"])

            # 발생시각 열 처리
            if "발생시각" in df.columns:
                if df["발생시각"].dtype == object:  # 문자열인 경우
                    # 시간 형식 변환 (HH:MM:SS 또는 HH:MM)
                    try:
                        df["발생시각_parsed"] = pd.to_datetime(
                            df["발생시각"], format="%H:%M:%S"
                        ).dt.time
                    except ValueError:
                        try:
                            df["발생시각_parsed"] = pd.to_datetime(
                                df["발생시각"], format="%H:%M"
                            ).dt.time
                        except ValueError:
                            df["발생시각_parsed"] = pd.to_datetime(
                                df["발생시각"], errors="coerce"
                            ).dt.time
                else:  # 이미 datetime인 경우
                    df["발생시각_parsed"] = df["발생시각"].dt.time

            return df
        except Exception as e:
            print(f"[ERROR] 날짜/시간 열 전처리 중 오류: {str(e)}")
            traceback.print_exc()
            return df

    def _convert_to_security_records(self, df: pd.DataFrame) -> List[SecurityRecord]:
        """
        데이터프레임을 SecurityRecord 객체 리스트로 변환합니다.

        Args:
            df (pd.DataFrame): 처리할 데이터프레임

        Returns:
            List[SecurityRecord]: SecurityRecord 객체 리스트
        """
        records = []
        for _, row in df.iterrows():
            try:
                # 날짜와 시간 가져오기
                record_date = (
                    row["발생일자_parsed"].date() if pd.notna(row["발생일자_parsed"]) else None
                )
                record_time = (
                    row["발생시각_parsed"]
                    if "발생시각_parsed" in row and pd.notna(row["발생시각_parsed"])
                    else None
                )

                if record_date is None or record_time is None:
                    continue

                # 모드 정보 가져오기
                mode = str(row["모드"]) if pd.notna(row["모드"]) else ""

                # 업무 일자 계산 (새벽 4시 기준)
                business_date = calculate_business_date(record_date, record_time)

                # 기록 유형 분류
                record_type = self._classify_record_type(mode)

                # SecurityRecord 객체 생성
                record = SecurityRecord(
                    record_date=record_date,
                    record_time=record_time,
                    mode=mode,
                    business_date=business_date,
                    record_type=record_type,
                )
                records.append(record)

            except Exception as e:
                print(f"[ERROR] 레코드 변환 중 오류: {str(e)}")
                continue

        return records

    def _classify_record_type(self, mode: str) -> str:
        """
        모드 값을 기반으로 기록 유형을 분류합니다.

        Args:
            mode (str): 모드 값

        Returns:
            str: 기록 유형 ("경비해제", "경비시작", "출입", "기타")
        """
        mode = str(mode).lower() if mode else ""

        # 경비 해제
        if any(keyword in mode for keyword in ["해제", "disarm", "off", "오픈"]):
            return "경비해제"

        # 경비 시작
        if any(
            keyword in mode for keyword in ["설정", "세팅", "arm", "on", "시작", "닫힘", "종료"]
        ):
            return "경비시작"

        # 출입 (명확하지 않음)
        if any(keyword in mode for keyword in ["출입", "입장", "퇴장", "access", "enter", "exit"]):
            return "출입"

        # 기타
        return "기타"

    def _track_security_status_changes(
        self, day_records: List[SecurityRecord]
    ) -> List[SecurityStatusChange]:
        """
        하루동안의 경비 상태 변화를 추적합니다.

        Args:
            day_records (List[SecurityRecord]): 하루 동안의 경비 기록 목록

        Returns:
            List[SecurityStatusChange]: 경비 상태 변화 목록
        """
        status_changes = []

        for record in day_records:
            if record.record_type == "경비해제":
                change = SecurityStatusChange(
                    time=record.record_time,
                    status="해제",
                    record=record,
                )
                status_changes.append(change)
            elif record.record_type == "경비시작":
                change = SecurityStatusChange(
                    time=record.record_time,
                    status="설정",
                    record=record,
                )
                status_changes.append(change)

        return status_changes

    def _derive_security_periods(
        self, status_changes: List[SecurityStatusChange]
    ) -> List[SecurityPeriod]:
        """
        경비 상태 변화로부터 경비 해제/설정 기간을 도출합니다.

        Args:
            status_changes (List[SecurityStatusChange]): 경비 상태 변화 목록

        Returns:
            List[SecurityPeriod]: 경비 기간 목록
        """
        if not status_changes:
            return []

        # 시간순으로 정렬
        status_changes.sort(key=lambda x: x.time)

        periods = []
        current_status = None
        start_time = None

        for change in status_changes:
            # 새로운 상태 시작
            if current_status is None:
                current_status = change.status
                start_time = change.time
                continue

            # 상태 변경
            if change.status != current_status:
                # 이전 기간 종료
                period = SecurityPeriod(
                    status=current_status,
                    start_time=start_time,
                    end_time=change.time,
                )
                periods.append(period)

                # 새 기간 시작
                current_status = change.status
                start_time = change.time

        # 마지막 변경 이후의 기간 (종료 시간은 모름)
        if current_status is not None and start_time is not None:
            period = SecurityPeriod(
                status=current_status,
                start_time=start_time,
                end_time=None,  # 종료 시간 불명확
            )
            periods.append(period)

        return periods
