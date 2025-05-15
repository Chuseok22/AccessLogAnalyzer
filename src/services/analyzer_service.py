"""
분석 서비스

경비 로그와 초과근무 로그를 비교 분석하는 핵심 서비스를 제공합니다.
"""

from typing import Dict, List, Any, Optional, Tuple
from datetime import datetime, time, date, timedelta
import traceback

from ..models.data_models import SuspiciousRecord


class AnalyzerService:
    """분석 서비스 클래스"""

    def __init__(self):
        self.no_security_records = []  # 경비 기록이 없는 날짜의 초과근무 기록

    def compare_security_and_overtime(
        self,
        security_records: Dict[date, List],
        overtime_records: List,
    ) -> List[SuspiciousRecord]:
        """
        경비 기록과 초과근무 기록을 비교 분석하여 의심스러운 기록을 반환합니다.

        1. 경비 기록이 없는 날짜에 초과근무 기록이 있는 경우
        2. 초과근무 시작 전에 경비 해제 기록이 없는 경우
        3. 초과근무 종료 후에 경비 시작 기록이 없는 경우

        Args:
            security_records (Dict): 날짜별 경비 기록 사전
            overtime_records (List): 초과근무 기록 목록

        Returns:
            List[SuspiciousRecord]: 의심스러운 초과근무 기록 목록
        """
        suspicious_records = []

        # 초과근무 기록이 없는 경우
        if not overtime_records:
            return suspicious_records

        # 경비 기록이 없는 경우
        if not security_records:
            # 모든 초과근무 기록을 의심 기록으로 추가
            for overtime in overtime_records:
                record = SuspiciousRecord(
                    record_date=overtime.date,
                    employee_name=overtime.name,
                    department=overtime.department,
                    overtime_hours=overtime.hours,
                    security_status="경비 기록 없음",
                    suspicious_reason="해당 날짜에 경비 기록이 전혀 없습니다.",
                    work_content=overtime.work_description,
                    is_holiday=overtime.is_holiday,
                )
                suspicious_records.append(record)
            return suspicious_records

        # 각 초과근무 기록에 대해 검사
        for overtime in overtime_records:
            overtime_date = overtime.date
            employee_name = overtime.name
            department = overtime.department
            start_time = overtime.start_time
            end_time = overtime.end_time
            work_description = overtime.work_description
            is_holiday = overtime.is_holiday

            # 경비 기록이 없는 날짜의 초과근무 기록
            if overtime_date not in security_records:
                record = SuspiciousRecord(
                    record_date=overtime_date,
                    employee_name=employee_name,
                    department=department,
                    overtime_hours=f"{start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}",
                    security_status="경비 기록 없음",
                    suspicious_reason="해당 날짜에 경비 기록이 전혀 없습니다.",
                    work_content=work_description,
                    is_holiday=is_holiday,
                )
                suspicious_records.append(record)
                self.no_security_records.append(overtime_date)  # 경비 기록 없는 날짜 추가
                continue

            # 해당 날짜의 경비 기록
            security_data = security_records[overtime_date]

            # 경비 기록이 비어있는 경우
            if not security_data or "periods" not in security_data:
                record = SuspiciousRecord(
                    record_date=overtime_date,
                    employee_name=employee_name,
                    department=department,
                    overtime_hours=f"{start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}",
                    security_status="경비 데이터 불완전",
                    suspicious_reason="해당 날짜의 경비 기록이 불완전합니다.",
                    work_content=work_description,
                    is_holiday=is_holiday,
                )
                suspicious_records.append(record)
                continue

            # 경비 해제 기간 확인
            security_periods = security_data["periods"]
            disarm_periods = []

            # 경비 해제 기간 목록 추출
            for period in security_periods:
                if period.status == "해제":
                    disarm_periods.append(period)

            # 경비 해제 기간이 없는 경우
            if not disarm_periods:
                record = SuspiciousRecord(
                    record_date=overtime_date,
                    employee_name=employee_name,
                    department=department,
                    overtime_hours=f"{start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}",
                    security_status="경비 해제 기록 없음",
                    suspicious_reason="해당 날짜에 경비 해제 기록이 없습니다.",
                    work_content=work_description,
                    is_holiday=is_holiday,
                )
                suspicious_records.append(record)
                continue

            # 초과근무 시간이 경비 해제 기간과 일치하는지 확인
            is_covered = False
            partially_covered = False
            partial_reason = ""

            for period in disarm_periods:
                disarm_start = period.start_time
                disarm_end = period.end_time

                # 경비 해제 시작/종료 시간이 None인 경우 처리
                if disarm_start is None or disarm_end is None:
                    continue

                # 초과근무 시간이 경비 해제 기간에 완전히 포함되는 경우
                if disarm_start <= start_time and end_time <= disarm_end:
                    is_covered = True
                    break

                # 초과근무 시작 시간만 경비 해제 기간에 포함되는 경우
                elif disarm_start <= start_time < disarm_end and end_time > disarm_end:
                    partially_covered = True
                    partial_reason = (
                        f"초과근무 종료({end_time.strftime('%H:%M')})가 "
                        f"경비 해제 종료({disarm_end.strftime('%H:%M')})보다 늦습니다."
                    )
                    break

                # 초과근무 종료 시간만 경비 해제 기간에 포함되는 경우
                elif start_time < disarm_start and disarm_start < end_time <= disarm_end:
                    partially_covered = True
                    partial_reason = (
                        f"초과근무 시작({start_time.strftime('%H:%M')})이 "
                        f"경비 해제 시작({disarm_start.strftime('%H:%M')})보다 빠릅니다."
                    )
                    break

            # 초과근무 시간이 경비 해제 기간과 일치하지 않는 경우
            if not is_covered:
                if partially_covered:
                    reason = f"초과근무 시간이 경비 해제 기간과 부분적으로만 일치합니다. {partial_reason}"
                else:
                    reason = "초과근무 시간이 경비 해제 기간과 일치하지 않습니다."

                record = SuspiciousRecord(
                    record_date=overtime_date,
                    employee_name=employee_name,
                    department=department,
                    overtime_hours=f"{start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}",
                    security_status="시간 불일치",
                    suspicious_reason=reason,
                    work_content=work_description,
                    is_holiday=is_holiday,
                )
                suspicious_records.append(record)

        return suspicious_records
