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


class OvertimeProcessor:
    """초과근무 기록 처리 클래스"""

    def __init__(self):
        self.missing_time_records = []  # 시간 정보가 누락된 기록
        self.error_records = []  # 처리 중 오류가 발생한 기록

    def process_overtime_log(
        self,
        df: pd.DataFrame,
        start_date: str = None,
        end_date: str = None,
    ) -> List[OvertimeRecord]:
        """
        초과근무 기록 엑셀 파일을 처리하여 초과근무 정보를 추출합니다.

        Args:
            df (pd.DataFrame): 초과근무 기록 데이터프레임 (헤더 없음으로 로드됨)
            start_date (str): 분석 시작 날짜 (YYYY-MM-DD)
            end_date (str): 분석 종료 날짜 (YYYY-MM-DD)

        Returns:
            List[OvertimeRecord]: 초과근무 기록 객체 목록
        """
        try:
            print("[DEBUG] 초과근무 기록 처리 시작")

            # 초기화
            self.missing_time_records.clear()
            self.error_records.clear()
            results = []

            # 데이터프레임이 비어있는 경우
            if df.empty:
                print("[WARNING] 초과근무 기록 데이터가 비어있습니다.")
                return results

            # 첫 2행은 헤더라고 가정
            if len(df) > 2:
                df = df.iloc[2:].reset_index(drop=True)
            else:
                print("[WARNING] 데이터가 충분하지 않습니다.")
                return results

            # 분석할 날짜 범위 설정
            filter_start = datetime.strptime(start_date, "%Y-%m-%d").date() if start_date else None
            filter_end = datetime.strptime(end_date, "%Y-%m-%d").date() if end_date else None

            # 각 행 처리
            for idx, row in df.iterrows():
                try:
                    # G열: 초과근무일자, H열: 출근시간, I열: 퇴근시간
                    overtime_date_col = 6  # G열 (0부터 시작)
                    start_time_col = 7  # H열
                    end_time_col = 8  # I열

                    # 날짜 처리
                    date_val = row.iloc[overtime_date_col]
                    if pd.isna(date_val):
                        continue

                    overtime_date = self._parse_date(date_val)
                    if overtime_date is None:
                        continue

                    # 날짜 필터링
                    if (filter_start and overtime_date < filter_start) or (
                        filter_end and overtime_date > filter_end
                    ):
                        continue

                    # 시간 처리
                    start_time_val = row.iloc[start_time_col] if start_time_col < len(row) else None
                    end_time_val = row.iloc[end_time_col] if end_time_col < len(row) else None

                    # 직원 정보 (C열: 사원번호, B열: 이름, E열: 부서)
                    emp_id = (
                        str(row.iloc[2]) if 2 < len(row) and not pd.isna(row.iloc[2]) else "N/A"
                    )
                    name = str(row.iloc[1]) if 1 < len(row) and not pd.isna(row.iloc[1]) else "N/A"
                    dept = str(row.iloc[4]) if 4 < len(row) and not pd.isna(row.iloc[4]) else "N/A"

                    # 근무 내용 (L열)
                    work_desc = (
                        str(row.iloc[11])
                        if 11 < len(row) and not pd.isna(row.iloc[11])
                        else "내용 미입력"
                    )

                    # 휴일여부 (K열)
                    holiday_val = (
                        row.iloc[10] if 10 < len(row) and not pd.isna(row.iloc[10]) else ""
                    )
                    is_holiday = "휴일" in str(holiday_val)

                    # 각 시간 값을 파싱
                    start_time = parse_time(start_time_val)
                    end_time = parse_time(end_time_val)

                    # 시간 정보 검증
                    if start_time is None or end_time is None:
                        # 시간 정보 누락 기록
                        self.missing_time_records.append(
                            {
                                "date": overtime_date,
                                "name": name,
                                "raw_start": start_time_val,
                                "raw_end": end_time_val,
                            }
                        )
                        continue

                    # 종료 시간이 시작 시간보다 이른 경우, 다음 날로 간주
                    hours_worked = self._calculate_hours(start_time, end_time)

                    # OvertimeRecord 객체 생성
                    record = OvertimeRecord(
                        date=overtime_date,
                        start_time=start_time,
                        end_time=end_time,
                        hours=hours_worked,
                        employee_id=emp_id,
                        name=name,
                        department=dept,
                        work_description=work_desc,
                        is_holiday=is_holiday,
                    )
                    results.append(record)

                except Exception as e:
                    # 행 처리 중 오류 발생
                    self.error_records.append((idx, str(e)))
                    print(f"[ERROR] 행 {idx} 처리 중 오류: {str(e)}")
                    traceback.print_exc()
                    continue

            print(f"[DEBUG] 초과근무 기록 처리 완료: {len(results)} 개 기록 처리됨")
            if self.missing_time_records:
                print(f"[WARNING] {len(self.missing_time_records)}개의 시간 정보 누락")
            if self.error_records:
                print(f"[WARNING] {len(self.error_records)}개의 처리 오류")

            return results

        except Exception as e:
            print(f"[ERROR] 초과근무 기록 처리 중 오류 발생: {str(e)}")
            traceback.print_exc()
            return []

    def _parse_date(self, date_val) -> Optional[date]:
        """
        다양한 형식의 날짜 값을 파싱하여 date 객체로 변환합니다.

        Args:
            date_val: 파싱할 날짜 값

        Returns:
            Optional[date]: 파싱된 날짜 객체 또는 None
        """
        if pd.isna(date_val):
            return None

        try:
            # 이미 날짜 객체인 경우
            if isinstance(date_val, (datetime, date)):
                return date_val.date() if isinstance(date_val, datetime) else date_val

            # 문자열인 경우
            if isinstance(date_val, str):
                # YYYY/MM/DD 또는 YYYY-MM-DD 형식 시도
                try:
                    parsed_date = datetime.strptime(date_val, "%Y/%m/%d").date()
                    return parsed_date
                except ValueError:
                    try:
                        parsed_date = datetime.strptime(date_val, "%Y-%m-%d").date()
                        return parsed_date
                    except ValueError:
                        pass

                # DD/MM/YYYY 또는 DD-MM-YYYY 형식 시도
                try:
                    parsed_date = datetime.strptime(date_val, "%d/%m/%Y").date()
                    return parsed_date
                except ValueError:
                    try:
                        parsed_date = datetime.strptime(date_val, "%d-%m-%Y").date()
                        return parsed_date
                    except ValueError:
                        pass

                # 한국어 날짜 형식 시도 (YYYY년 MM월 DD일)
                try:
                    date_val = date_val.replace(" ", "")  # 공백 제거
                    date_val = (
                        date_val.replace("년", "-")
                        .replace("월", "-")
                        .replace("일", "")
                        .replace(".", "-")
                    )
                    parsed_date = datetime.strptime(date_val, "%Y-%m-%d").date()
                    return parsed_date
                except ValueError:
                    pass

            # 숫자 형식인 경우 (엑셀 날짜 값)
            try:
                parsed_date = pd.Timestamp.fromordinal(
                    datetime(1900, 1, 1).toordinal() + int(date_val) - 2
                ).date()
                return parsed_date
            except (ValueError, TypeError, OverflowError):
                pass

            # pandas를 사용한 일반적인 변환 시도
            return pd.to_datetime(date_val).date()

        except Exception as e:
            print(f"[ERROR] 날짜 파싱 오류 ({date_val}): {str(e)}")
            return None

    def _calculate_hours(self, start_time: time, end_time: time) -> str:
        """
        시작 시간과 종료 시간으로부터 근무 시간을 계산합니다.
        종료 시간이 시작 시간보다 이른 경우, 자정을 넘어간 것으로 간주합니다.

        Args:
            start_time (time): 시작 시간
            end_time (time): 종료 시간

        Returns:
            str: "HH:MM - HH:MM" 형식의 근무 시간 문자열
        """
        return f"{start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}"
