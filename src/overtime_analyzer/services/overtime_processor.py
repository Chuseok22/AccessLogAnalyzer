"""
초과근무 기록 처리 서비스

초과근무 기록 엑셀 파일을 처리하여 초과근무 정보를 분석하는 기능을 제공합니다.
"""

import pandas as pd
import traceback
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime, time, date, timedelta

from ..models.data_models import OvertimeRecord
from ..utils.date_utils import parse_time, is_valid_time_format, calculate_business_date


class OvertimeLogProcessor:
    """초과근무 기록 처리 클래스"""

    def __init__(self):
        self.missing_time_records = []  # 시간 정보가 누락된 기록
        self.error_records = []  # 처리 중 오류가 발생한 기록

    def process_overtime_log(
        self, df: pd.DataFrame, start_date: str, end_date: str
    ) -> List[Dict[str, Any]]:
        """
        초과근무 기록을 처리합니다.

        Args:
            df: 초과근무 기록 데이터프레임
            start_date: 시작 날짜 (YYYY-MM-DD 형식)
            end_date: 종료 날짜 (YYYY-MM-DD 형식)

        Returns:
            List[Dict[str, Any]]: 처리된 초과근무 기록 목록
        """
        try:
            # 데이터가 3행부터 시작하고, 열 위치가 고정된 경우를 처리
            df = self._preprocess_dataframe(df)

            # 위치 기반으로 컬럼 매핑
            col_mapping = self._map_columns(df)

            # 날짜 데이터 처리
            df = self._process_date_column(df, col_mapping)

            # 필터링 적용
            filtered_df = self._apply_date_filter(df, col_mapping, start_date, end_date)

            # 각 행의 데이터 유효성 확인
            filtered_df = self._validate_data(filtered_df, col_mapping)

            # 초과근무 데이터 처리
            overtime_records = self._process_overtime_records(filtered_df, col_mapping)

            # 시간 누락 기록 확인
            self._check_missing_time_records(filtered_df, col_mapping)

            return overtime_records

        except Exception as e:
            # 예외 발생 시 상세 정보 출력하고 다시 발생
            print(f"초과근무 기록 처리 중 오류: {str(e)}")
            print(traceback.format_exc())
            raise

    def _preprocess_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        데이터프레임 전처리 - 3행부터 시작하는 데이터 처리

        Args:
            df: 원본 데이터프레임

        Returns:
            pd.DataFrame: 전처리된 데이터프레임
        """
        # 데이터가 3행부터 시작하고, 열 위치가 고정된 경우를 처리
        if len(df) >= 2:
            # 필요하면 1-2행 제거 (헤더 및 SQL 실행 결과)
            df = df.iloc[2:].reset_index(drop=True)
        return df

    def _map_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        컬럼 매핑 - 위치 기반으로 컬럼을 매핑합니다.

        Args:
            df: 처리할 데이터프레임

        Returns:
            Dict[str, str]: 컬럼 매핑 정보
        """
        # 인덱스 기반으로 컬럼 이름 만들기
        renamed_columns = {}
        for i in range(len(df.columns)):
            if i == 0:
                renamed_columns[df.columns[i]] = "부서명"
            elif i == 1:
                renamed_columns[df.columns[i]] = "직급"
            elif i == 2:
                renamed_columns[df.columns[i]] = "개인식별번호"
            elif i == 3:
                renamed_columns[df.columns[i]] = "성명"
            elif i == 4:
                renamed_columns[df.columns[i]] = "현업여부"
            elif i == 5:
                renamed_columns[df.columns[i]] = "휴일여부"
            elif i == 6:
                renamed_columns[df.columns[i]] = "초과근무일자"
            elif i == 7:
                renamed_columns[df.columns[i]] = "출근시간"
            elif i == 8:
                renamed_columns[df.columns[i]] = "퇴근시간"
            elif i == 9:
                renamed_columns[df.columns[i]] = "출근IP"
            elif i == 10:
                renamed_columns[df.columns[i]] = "퇴근IP"
            elif i == 11:
                renamed_columns[df.columns[i]] = "초과근무시간"
            elif i == 12:
                renamed_columns[df.columns[i]] = "수당시간"
            elif i == 13:
                renamed_columns[df.columns[i]] = "근무내용"
            else:
                renamed_columns[df.columns[i]] = f"컬럼{i}"

        # 데이터프레임 컬럼 이름 변경
        df = df.rename(columns=renamed_columns, inplace=True)

        # 컬럼 이름 기반으로 매핑
        col_mapping = {
            "날짜": "초과근무일자",
            "시작시간": "출근시간",
            "종료시간": "퇴근시간",
            "이름": "성명",
        }

        return col_mapping

    def _process_date_column(self, df: pd.DataFrame, col_mapping: Dict[str, str]) -> pd.DataFrame:
        """
        날짜 컬럼을 처리합니다.

        Args:
            df: 처리할 데이터프레임
            col_mapping: 컬럼 매핑 정보

        Returns:
            pd.DataFrame: 처리된 데이터프레임
        """
        # 날짜 데이터 정리 (YYYY-MM-DD 형식 고정)
        # 데이터프레임에서 초과근무일자 열이 문자열이면 datetime으로 변환
        df["날짜_datetime"] = pd.to_datetime(
            df[col_mapping["날짜"]], format="%Y-%m-%d", errors="coerce"
        )

        return df

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
            start_date_obj = datetime.strptime(start_date, "%Y-%m-%d").date()
            end_date_obj = datetime.strptime(end_date, "%Y-%m-%d").date()

            filtered_df = df[
                (df["날짜_datetime"] >= pd.Timestamp(start_date_obj))
                & (df["날짜_datetime"] <= pd.Timestamp(end_date_obj))
            ]
            return filtered_df
        return df

    def _validate_data(self, df: pd.DataFrame, col_mapping: Dict[str, str]) -> pd.DataFrame:
        """
        데이터의 유효성을 검사합니다.

        Args:
            df: 처리할 데이터프레임
            col_mapping: 컬럼 매핑 정보

        Returns:
            pd.DataFrame: 유효한 데이터만 포함된 데이터프레임
        """
        # 각 행의 데이터 유효성 확인을 위한 작업
        df["데이터_유효"] = (
            pd.notna(df[col_mapping["이름"]])
            & pd.notna(df[col_mapping["날짜"]])
            & pd.notna(df[col_mapping["시작시간"]])
            & pd.notna(df[col_mapping["종료시간"]])
            & pd.notna(df["날짜_datetime"])
        )

        # 유효한 데이터만 선택
        return df[df["데이터_유효"]]

    def _process_overtime_records(
        self, filtered_df: pd.DataFrame, col_mapping: Dict[str, str]
    ) -> List[Dict[str, Any]]:
        """
        초과근무 기록을 처리합니다.

        Args:
            filtered_df: 필터링된 데이터프레임
            col_mapping: 컬럼 매핑 정보

        Returns:
            List[Dict[str, Any]]: 처리된 초과근무 기록 목록
        """
        overtime_records = []

        # 정규 근무시간 설정 (9:00-18:00)
        regular_start = time(9, 0)
        regular_end = time(18, 0)

        for _, row in filtered_df.iterrows():
            try:
                # 날짜 처리
                work_date = row["날짜_datetime"].date()

                # 출/퇴근 시간 누락 여부 플래그
                has_missing_time = False
                missing_time_fields = []

                # 출근시간 처리 (HH:mm 고정 형식)
                if pd.isna(row[col_mapping["시작시간"]]):
                    has_missing_time = True
                    missing_time_fields.append("출근시간")

                start_time = parse_time(row[col_mapping["시작시간"]], regular_start)

                if start_time is None:
                    # 시작 시간이 없거나 파싱 실패
                    has_missing_time = True
                    missing_time_fields.append("출근시간")
                    start_time = regular_start  # 일단 기본값 설정 (나중에 의심 데이터로 표시)

                # 퇴근시간 처리 (HH:mm 고정 형식)
                if pd.isna(row[col_mapping["종료시간"]]):
                    has_missing_time = True
                    missing_time_fields.append("퇴근시간")

                end_time = parse_time(row[col_mapping["종료시간"]], regular_end)

                if end_time is None:
                    # 종료 시간이 없거나 파싱 실패
                    has_missing_time = True
                    missing_time_fields.append("퇴근시간")
                    end_time = regular_end  # 일단 기본값 설정 (나중에 의심 데이터로 표시)

                # 업무일 결정 (새벽 4시 기준)
                business_date = work_date

                # 자정 이후 새벽 4시 이전 근무는 전날 업무일로 계산
                if end_time < time(4, 0):
                    business_date = work_date - timedelta(days=1)

                # 휴일 여부 확인 (F열 데이터만 사용)
                is_holiday = self._check_if_holiday(row)

                # 직원 이름 정보
                employee_name = (
                    str(row[col_mapping["이름"]])
                    if pd.notna(row[col_mapping["이름"]])
                    else "Unknown"
                )

                # 부서명 정보 추가 (있는 경우)
                department = (
                    str(row["부서명"]) if "부서명" in row and pd.notna(row["부서명"]) else ""
                )

                # 초과근무시간 정보 (있는 경우 사용)
                overtime_hours = self._parse_overtime_hours(row)

                # 추가 정보 (근무내용)
                work_description = ""
                if "근무내용" in row and pd.notna(row["근무내용"]):
                    work_description = str(row["근무내용"]).strip()

                # 휴일 또는 초과근무 정보 처리
                self._add_overtime_record(
                    overtime_records,
                    business_date,
                    work_date,
                    start_time,
                    end_time,
                    is_holiday,
                    employee_name,
                    department,
                    overtime_hours,
                    work_description,
                    regular_start,
                    regular_end,
                )

            except Exception as e:
                print(f"초과근무 기록 처리 중 오류: {str(e)}")
                error_desc = str(e)

                # 오류 발생 데이터 저장
                # 유효한 값들은 try 블록 내에서 정의되었으므로
                # 여기서는 현재 스코프에 있는 변수만 전달
                try:
                    # business_date와 employee_name이 이 스코프에 정의되어 있을 수 있음
                    self._add_error_record(
                        row,
                        col_mapping,
                        error_desc,
                        business_date=locals().get("business_date"),
                        work_date=locals().get("work_date"),
                        employee_name=locals().get("employee_name"),
                    )
                except NameError:
                    # 변수가 정의되지 않은 경우
                    self._add_error_record(row, col_mapping, error_desc)

        return overtime_records

    def _check_if_holiday(self, row: pd.Series) -> bool:
        """
        휴일 여부를 확인합니다.

        Args:
            row: 데이터 행

        Returns:
            bool: 휴일 여부
        """
        is_holiday = False

        # F열(휴일여부) 데이터 확인
        if "휴일여부" in row and pd.notna(row["휴일여부"]):
            holiday_value = str(row["휴일여부"]).strip().lower()
            # "Y", "휴일", "공휴일", "토요일", "일요일" 등의 문자가 포함되어 있으면 휴일로 판단
            holiday_keywords = ["y", "휴", "공휴", "토요일", "일요일"]
            is_holiday = any(x in holiday_value for x in holiday_keywords)
            matched_keywords = [x for x in holiday_keywords if x in holiday_value]

            if matched_keywords:
                print(f"[휴일감지] 휴일키워드 '{matched_keywords[0]}' 감지됨: '{holiday_value}'")
            else:
                print(f"[평일감지] 휴일키워드 없음: '{holiday_value}'")
        else:
            # F열 데이터가 없거나 누락된 경우 기본값은 평일(False)로 설정
            print(
                f"[데이터없음] 평일로 처리(기본값), F열 데이터: '{row.get('휴일여부', '데이터 없음')}'"
            )

        return is_holiday

    def _parse_overtime_hours(self, row: pd.Series) -> Optional[float]:
        """
        초과근무시간을 파싱합니다.

        Args:
            row: 데이터 행

        Returns:
            Optional[float]: 파싱된 초과근무시간 (시간) 또는 None
        """
        overtime_hours = None
        if "초과근무시간" in row and pd.notna(row["초과근무시간"]):
            try:
                if isinstance(row["초과근무시간"], str):
                    # 문자열 형식 처리 ("1:30" -> 1.5시간)
                    time_str = row["초과근무시간"].replace(" ", "")
                    if ":" in time_str:
                        hours, minutes = map(int, time_str.split(":"))
                        overtime_hours = hours + minutes / 60.0
                    else:
                        # 숫자 형식 문자열인 경우
                        overtime_hours = float(time_str)
                else:
                    # 숫자 형식인 경우 직접 변환
                    overtime_hours = float(row["초과근무시간"])
            except ValueError as parse_err:
                print(f"초과근무시간 변환 실패: {row['초과근무시간']} ({parse_err})")

        return overtime_hours

    def _add_overtime_record(
        self,
        overtime_records: List[Dict[str, Any]],
        business_date: date,
        work_date: date,
        start_time: time,
        end_time: time,
        is_holiday: bool,
        employee_name: str,
        department: str,
        overtime_hours: Optional[float],
        work_description: str,
        regular_start: time,
        regular_end: time,
    ) -> None:
        """
        초과근무 기록을 추가합니다.

        Args:
            overtime_records: 초과근무 기록 목록
            business_date: 업무일
            work_date: 원래 날짜
            start_time: 시작 시간
            end_time: 종료 시간
            is_holiday: 휴일 여부
            employee_name: 직원명
            department: 부서명
            overtime_hours: 기록된 초과근무 시간
            work_description: 근무 내용 설명
            regular_start: 정규 근무 시작 시간
            regular_end: 정규 근무 종료 시간
        """
        # 휴일인 경우 모든 시간을 초과근무로 처리
        if is_holiday:
            # 휴일 근무는 전체가 초과근무
            overtime_start = start_time
            overtime_end = end_time

            overtime_records.append(
                {
                    "업무일": business_date,
                    "날짜": work_date,
                    "시작시간": overtime_start,
                    "종료시간": overtime_end,
                    "초과근무유형": "휴일근무",
                    "직원명": employee_name,
                    "부서명": department,
                    "기록된_초과근무시간": overtime_hours,
                    "근무내용": work_description,
                    "휴일여부": True,
                }
            )

            # 디버그 정보 출력
            print(f"[휴일처리] {business_date} - {employee_name} - 휴일로 처리됨")

        # 평일인 경우 정규 근무시간(9-18)을 제외한 시간만 초과근무로 처리
        else:
            # 초과근무 여부 확인
            is_overtime = start_time >= regular_end or end_time <= regular_start

            # 업무 시간이 정규 근무시간(9-18)에 걸쳐있는 경우, 그 부분은 초과근무가 아님
            if not is_overtime and (start_time < regular_end and end_time > regular_start):
                # 시작 시간이 9시 이전이면 초과근무 시작 부분 기록
                if start_time < regular_start:
                    overtime_start = start_time
                    overtime_end = regular_start
                    overtime_type = "조기출근"

                    overtime_records.append(
                        {
                            "업무일": business_date,
                            "날짜": work_date,
                            "시작시간": overtime_start,
                            "종료시간": overtime_end,
                            "초과근무유형": overtime_type,
                            "직원명": employee_name,
                            "부서명": department,
                            "기록된_초과근무시간": overtime_hours,
                            "근무내용": work_description,
                            "휴일여부": False,
                        }
                    )

                # 종료 시간이 18시 이후면 초과근무 종료 부분 기록
                if end_time > regular_end:
                    overtime_start = regular_end
                    overtime_end = end_time
                    overtime_type = "야간근무"

                    overtime_records.append(
                        {
                            "업무일": business_date,
                            "날짜": work_date,
                            "시작시간": overtime_start,
                            "종료시간": overtime_end,
                            "초과근무유형": overtime_type,
                            "직원명": employee_name,
                            "부서명": department,
                            "기록된_초과근무시간": overtime_hours,
                            "근무내용": work_description,
                            "휴일여부": False,
                        }
                    )

            elif is_overtime:
                # 전체가 초과근무인 경우
                overtime_start = start_time
                overtime_end = end_time

                # 시간대에 따른 초과근무 유형 결정
                if overtime_start < regular_start:
                    overtime_type = "조기출근"
                else:
                    overtime_type = "야간근무"

                overtime_records.append(
                    {
                        "업무일": business_date,
                        "날짜": work_date,
                        "시작시간": overtime_start,
                        "종료시간": overtime_end,
                        "초과근무유형": overtime_type,
                        "직원명": employee_name,
                        "부서명": department,
                        "기록된_초과근무시간": overtime_hours,
                        "근무내용": work_description,
                        "휴일여부": False,
                    }
                )

    def _add_error_record(
        self,
        row: pd.Series,
        col_mapping: Dict[str, str],
        error_desc: str,
        business_date: Optional[date] = None,
        work_date: Optional[date] = None,
        employee_name: Optional[str] = None,
    ) -> None:
        """
        오류 기록을 추가합니다.

        Args:
            row: 오류가 발생한 데이터 행
            col_mapping: 컬럼 매핑 정보
            error_desc: 오류 설명
            business_date: 업무일 (기본값: None)
            work_date: 실제 작업 날짜 (기본값: None)
            employee_name: 직원 이름 (기본값: None)
        """
        # 변수가 없는 경우 행에서 직원 이름 가져오기
        if employee_name is None and col_mapping["이름"] in row:
            employee_name = row.get(col_mapping["이름"], "알 수 없음")

        # 오류 발생 데이터 저장
        error_record = {
            "업무일": (
                business_date
                if business_date is not None
                else (work_date if work_date is not None else "알 수 없음")
            ),
            "직원명": employee_name or "알 수 없음",
            "오류내용": error_desc,
            "원본데이터": {
                k: str(v)
                for k, v in row.items()
                if k
                in [
                    col_mapping["날짜"],
                    col_mapping["시작시간"],
                    col_mapping["종료시간"],
                    col_mapping["이름"],
                    "부서명",
                    "근무내용",
                ]
            },
        }

        # 오류 데이터 목록에 추가
        self.error_records.append(error_record)

    def _check_missing_time_records(
        self, filtered_df: pd.DataFrame, col_mapping: Dict[str, str]
    ) -> None:
        """
        시간 정보가 누락된 기록을 확인합니다.

        Args:
            filtered_df: 처리할 데이터프레임
            col_mapping: 컬럼 매핑 정보
        """
        self.missing_time_records = []

        for record in filtered_df.iterrows():
            row = record[1]
            if pd.isna(row[col_mapping["시작시간"]]) or pd.isna(row[col_mapping["종료시간"]]):
                work_date = row["날짜_datetime"].date() if pd.notna(row["날짜_datetime"]) else None
                employee_name = (
                    str(row[col_mapping["이름"]])
                    if pd.notna(row[col_mapping["이름"]])
                    else "Unknown"
                )
                department = (
                    str(row["부서명"]) if "부서명" in row and pd.notna(row["부서명"]) else ""
                )

                missing_fields = []
                if pd.isna(row[col_mapping["시작시간"]]):
                    missing_fields.append("출근시간")
                if pd.isna(row[col_mapping["종료시간"]]):
                    missing_fields.append("퇴근시간")

                self.missing_time_records.append(
                    {
                        "업무일": work_date,
                        "직원명": employee_name,
                        "부서명": department,
                        "누락필드": ", ".join(missing_fields),
                        "원본데이터": {
                            k: str(v)
                            for k, v in row.items()
                            if k
                            in [
                                col_mapping["날짜"],
                                col_mapping["시작시간"],
                                col_mapping["종료시간"],
                                col_mapping["이름"],
                                "부서명",
                                "근무내용",
                            ]
                        },
                    }
                )

        # 시간 누락 기록 안내
        if self.missing_time_records:
            print(
                f"[주의] {len(self.missing_time_records)}개의 출/퇴근 시간 누락 기록이 발견되었습니다."
            )
