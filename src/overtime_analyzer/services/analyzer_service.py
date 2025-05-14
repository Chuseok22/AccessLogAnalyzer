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
        security_status_by_day: Dict[date, List[Dict[str, str]]],
        overtime_records: List[Dict[str, Any]],
    ) -> List[SuspiciousRecord]:
        """
        경비 상태와 초과근무 기록을 비교 분석하여 의심스러운 기록을 찾습니다.

        Args:
            security_status_by_day: 날짜별 경비 상태 변화 목록
            overtime_records: 초과근무 기록 목록

        Returns:
            List[SuspiciousRecord]: 의심스러운 초과근무 기록 목록
        """
        suspicious_records = []

        # 각 초과근무 기록에 대해 경비 상태 확인
        for overtime in overtime_records:
            business_date = overtime["업무일"]
            employee_name = overtime["직원명"]
            overtime_start = overtime["시작시간"]
            overtime_end = overtime["종료시간"]

            # 업무일에 해당하는 경비 기록 찾기
            security_status = security_status_by_day.get(business_date, [])

            # 경비 기록이 없는 경우 처리
            if not security_status:
                self._handle_no_security_record(
                    suspicious_records,
                    business_date,
                    employee_name,
                    overtime_start,
                    overtime_end,
                    overtime_records,
                )
                continue

            # 경비 기록이 있는 경우 분석
            result = self._analyze_security_overtime_overlap(
                business_date,
                employee_name,
                overtime_start,
                overtime_end,
                security_status,
                overtime_records,
            )

            # 의심 기록이 발견된 경우 추가
            if result:
                suspicious_records.append(result)

        return suspicious_records

    def _handle_no_security_record(
        self,
        suspicious_records: List[SuspiciousRecord],
        business_date: date,
        employee_name: str,
        overtime_start: time,
        overtime_end: time,
        overtime_records: List[Dict[str, Any]],
    ) -> None:
        """
        경비 기록이 없는 경우를 처리합니다.

        Args:
            suspicious_records: 의심 기록 목록
            business_date: 업무일
            employee_name: 직원명
            overtime_start: 초과근무 시작 시간
            overtime_end: 초과근무 종료 시간
            overtime_records: 초과근무 기록 목록
        """
        # 해당 업무일의 경비 기록이 없는 경우 의심 데이터로 저장
        if not hasattr(self, "no_security_records"):
            self.no_security_records = []

        self.no_security_records.append(
            {
                "업무일": business_date,
                "직원명": employee_name,
                "초과근무시간": f"{overtime_start.strftime('%H:%M')}-{overtime_end.strftime('%H:%M')}",
                "문제": "경비 기록 없음",
            }
        )
        print(f"[의심데이터] {business_date} - {employee_name} - 경비 기록 없음 (사용자 확인 필요)")

        # 경비 기록이 없어도 의심 데이터로 추가
        suspicious_reason = "해당 업무일에 경비 기록 없음"

        # 초과근무 기록에서 추가 정보 수집
        work_content = ""
        department = ""
        is_holiday = False

        for ovt_record in overtime_records:
            if ovt_record["직원명"] == employee_name and ovt_record["업무일"] == business_date:
                if "부서명" in ovt_record and ovt_record["부서명"]:
                    department = ovt_record["부서명"]
                if "근무내용" in ovt_record and ovt_record["근무내용"]:
                    work_content = ovt_record["근무내용"]
                if "휴일여부" in ovt_record:
                    is_holiday = ovt_record["휴일여부"]
                break

        overtime_period = f"{overtime_start.strftime('%H:%M')}-{overtime_end.strftime('%H:%M')}"

        # SuspiciousRecord 객체 생성
        suspicious_records.append(
            SuspiciousRecord(
                date=business_date,
                employee_name=employee_name,
                department=department,
                overtime_hours=overtime_period,
                overtime_period=overtime_period,
                security_status="기록 없음",
                suspicious_reason=suspicious_reason,
                work_content=work_content,
                is_holiday=is_holiday,
            )
        )

    def _analyze_security_overtime_overlap(
        self,
        business_date: date,
        employee_name: str,
        overtime_start: time,
        overtime_end: time,
        security_status: List[Dict[str, str]],
        overtime_records: List[Dict[str, Any]],
    ) -> Optional[SuspiciousRecord]:
        """
        경비 상태와 초과근무 시간이 겹치는지 분석합니다.

        Args:
            business_date: 업무일
            employee_name: 직원명
            overtime_start: 초과근무 시작 시간
            overtime_end: 초과근무 종료 시간
            security_status: 해당 업무일의 경비 상태 변화 목록
            overtime_records: 초과근무 기록 전체 목록

        Returns:
            Optional[SuspiciousRecord]: 의심 기록 정보 또는 None
        """
        try:
            # 초과근무 시간과 경비 상태 비교
            suspicious_reason = None

            # 초과근무 시간을 datetime으로 변환
            overtime_start_dt = datetime.combine(business_date, overtime_start)
            overtime_end_dt = datetime.combine(business_date, overtime_end)

            # 자정을 넘어가는 경우 다음날로 설정
            if overtime_end < overtime_start:
                overtime_end_dt = datetime.combine(business_date + timedelta(days=1), overtime_end)

            # 경비 기록을 시간순으로 정렬
            security_status.sort(key=lambda x: x["시간"])

            # 경비 상태 변화 기록
            security_changes = []

            # 경비 상태 기록을 시간 순으로 정렬하고 상태 변화 추적
            for record in security_status:
                record_time = datetime.strptime(record["시간"], "%Y-%m-%d %H:%M")
                security_changes.append({"시간": record_time, "상태": record["상태"]})

            # 경비 변화가 없으면 정상 종료
            if not security_changes:
                return None

            # 초기 상태 설정
            security_changes.sort(key=lambda x: x["시간"])

            # 의심 시간 구간 계산
            suspicious_intervals = []

            # 경비 상태의 시간대별 구간 생성 (경비시작-경비해제 구간)
            security_periods = self._create_security_periods(security_changes, business_date)

            # 경비기록이 하나도 없는 경우 (상태 변화가 없는 경우)
            if len(security_periods) == 0:
                # 기본적으로 보수적인 접근: 경비 활성화 상태로 간주
                today_start = datetime.combine(business_date, time(0, 0))
                tomorrow_start = datetime.combine(business_date + timedelta(days=1), time(0, 0))
                security_periods.append(
                    {
                        "시작": today_start,
                        "종료": tomorrow_start,
                        "상태": "시작",  # 경비 활성화 상태
                    }
                )

            # 초과근무 시간과 경비 활성화 시간대를 비교하여 의심 구간 계산
            suspicious_intervals = self._check_suspicious_intervals(
                overtime_start_dt, overtime_end_dt, security_periods, business_date, employee_name
            )

            # 모든 의심 구간에 대해 총 중첩 시간 계산
            total_suspicious_hours = 0
            suspicious_periods = []

            for start_time, end_time in suspicious_intervals:
                if start_time < end_time:  # 유효한 구간만 처리
                    duration_hours = (end_time - start_time).seconds / 3600
                    if duration_hours > 0:  # 1분이라도 중첩되면 의심 구간으로 간주
                        total_suspicious_hours += duration_hours
                        start_str = start_time.strftime("%H:%M")
                        end_str = end_time.strftime("%H:%M")
                        suspicious_periods.append(f"{start_str}-{end_str}")

            # 의심 시간이 있으면 의심 기록 생성
            if total_suspicious_hours > 0:
                return self._create_suspicious_record(
                    business_date,
                    employee_name,
                    overtime_start,
                    overtime_end,
                    total_suspicious_hours,
                    suspicious_periods,
                    security_changes,
                    overtime_start_dt,
                    overtime_end_dt,
                    overtime_records,
                )

            return None
        except Exception as e:
            print(f"경비-초과근무 비교 분석 중 오류: {str(e)}")
            print(traceback.format_exc())
            return None

    def _create_security_periods(
        self, security_changes: List[Dict[str, Any]], business_date: date
    ) -> List[Dict[str, Any]]:
        """
        경비 상태 변화 기록을 바탕으로 시간대별 구간을 생성합니다.

        Args:
            security_changes: 경비 상태 변화 기록
            business_date: 업무일

        Returns:
            List[Dict[str, Any]]: 경비 상태 시간대 구간 목록
        """
        security_periods = []

        if len(security_changes) > 0:
            # 먼저 첫 상태가 "해제"인 경우, 자정부터 첫 해제까지는 경비 활성화 상태로 간주
            if security_changes[0]["상태"] == "해제":
                midnight = datetime.combine(business_date, time(0, 0))
                security_periods.append(
                    {
                        "시작": midnight,
                        "종료": security_changes[0]["시간"],
                        "상태": "시작",  # 경비 활성화 상태
                    }
                )

            # 이후의 상태 변화를 추적하며 구간 생성
            for i in range(len(security_changes)):
                current = security_changes[i]

                # 마지막 항목이거나 다음 항목의 상태가 현재와 다른 경우
                if i == len(security_changes) - 1:
                    # 마지막 상태가 "시작"인 경우, 해당 시작부터 자정까지 경비 활성화 상태로 간주
                    if current["상태"] == "시작":
                        next_day = datetime.combine(business_date + timedelta(days=1), time(0, 0))
                        security_periods.append(
                            {"시작": current["시간"], "종료": next_day, "상태": "시작"}
                        )
                else:
                    next_change = security_changes[i + 1]
                    # 현재 상태부터 다음 상태 변경 전까지의 구간 생성
                    security_periods.append(
                        {
                            "시작": current["시간"],
                            "종료": next_change["시간"],
                            "상태": current["상태"],
                        }
                    )

        return security_periods

    def _check_suspicious_intervals(
        self,
        overtime_start_dt: datetime,
        overtime_end_dt: datetime,
        security_periods: List[Dict[str, Any]],
        business_date: date,
        employee_name: str,
    ) -> List[Tuple[datetime, datetime]]:
        """
        초과근무 시간과 경비 활성화 구간이 겹치는지 확인하여 의심 구간을 계산합니다.

        Args:
            overtime_start_dt: 초과근무 시작 datetime
            overtime_end_dt: 초과근무 종료 datetime
            security_periods: 경비 상태 시간대 구간 목록
            business_date: 업무일
            employee_name: 직원명

        Returns:
            List[Tuple[datetime, datetime]]: 의심 시간 구간 목록
        """
        suspicious_intervals = []

        # 각 경비 활성화 구간과 초과근무 시간 비교
        for period in security_periods:
            # 경비 활성화 상태인 경우만 검사
            if period["상태"] == "시작":
                # 초과근무 시간이 경비 활성화 구간과 겹치는지 확인
                if max(period["시작"], overtime_start_dt) < min(period["종료"], overtime_end_dt):
                    # 겹치는 구간 계산
                    overlap_start = max(period["시작"], overtime_start_dt)
                    overlap_end = min(period["종료"], overtime_end_dt)
                    suspicious_intervals.append((overlap_start, overlap_end))

                    # 디버그 출력
                    print(
                        f"[의심기록] {business_date} - {employee_name} - "
                        f"경비활성화({period['시작'].strftime('%H:%M:%S')}-"
                        f"{period['종료'].strftime('%H:%M:%S')}) 중 "
                        f"초과근무 발생({overlap_start.strftime('%H:%M:%S')}-"
                        f"{overlap_end.strftime('%H:%M:%S')})"
                    )

        return suspicious_intervals

    def _create_suspicious_record(
        self,
        business_date: date,
        employee_name: str,
        overtime_start: time,
        overtime_end: time,
        total_suspicious_hours: float,
        suspicious_periods: List[str],
        security_changes: List[Dict[str, Any]],
        overtime_start_dt: datetime,
        overtime_end_dt: datetime,
        overtime_records: List[Dict[str, Any]],
    ) -> SuspiciousRecord:
        """
        의심스러운 초과근무 기록을 생성합니다.

        Args:
            business_date: 업무일
            employee_name: 직원명
            overtime_start: 초과근무 시작 시간
            overtime_end: 초과근무 종료 시간
            total_suspicious_hours: 총 의심 시간 (시간)
            suspicious_periods: 의심 시간대 문자열 목록
            security_changes: 경비 상태 변화 기록
            overtime_start_dt: 초과근무 시작 datetime
            overtime_end_dt: 초과근무 종료 datetime
            overtime_records: 초과근무 기록 전체 목록

        Returns:
            SuspiciousRecord: 의심 기록 객체
        """
        period_str = ", ".join(suspicious_periods)
        suspicious_reason = (
            f"경비 작동 중 총 {total_suspicious_hours:.1f}시간 초과근무 기록 존재 ({period_str})"
        )

        # 디버그 정보: 의심 세부 정보
        print(f"[의심결과] {business_date} - {employee_name}")
        print(
            f"  ├─ 초과근무: {overtime_start_dt.strftime('%H:%M')}-{overtime_end_dt.strftime('%H:%M')}"
        )
        print(f"  ├─ 의심구간: {period_str}")
        print(f"  └─ 총 의심시간: {total_suspicious_hours:.2f}시간")

        # 초과근무 기록에서 추가 정보 찾기
        department = ""
        work_content = ""
        is_holiday = False

        # 해당 직원의 초과근무 기록 중에서 부가 정보 찾기
        for ovt_record in overtime_records:
            if ovt_record["직원명"] == employee_name and ovt_record["업무일"] == business_date:
                if "부서명" in ovt_record and ovt_record["부서명"]:
                    department = ovt_record["부서명"]
                if "근무내용" in ovt_record and ovt_record["근무내용"]:
                    work_content = ovt_record["근무내용"]
                if "휴일여부" in ovt_record:
                    is_holiday = ovt_record["휴일여부"]
                break

        # 휴일 여부에 따라 의심 사유 보완
        if is_holiday:
            suspicious_reason += " (휴일 근무)"

        # 결과에 추가할 의심 정보 준비
        security_info = "경비 작동 중"

        # 경비 설정 시간 정보가 있으면 포함
        security_set_times = []
        for change in security_changes:
            if change["상태"] == "시작" and overtime_start_dt <= change["시간"] <= overtime_end_dt:
                security_set_times.append(change["시간"].strftime("%H:%M:%S"))

        if security_set_times:
            security_info += f" (경비설정시각: {', '.join(security_set_times)})"

        # 초과근무 시간 표시용 문자열
        overtime_period = f"{overtime_start.strftime('%H:%M')}-{overtime_end.strftime('%H:%M')}"

        # SuspiciousRecord 객체 생성
        return SuspiciousRecord(
            date=business_date,
            employee_name=employee_name,
            department=department,
            overtime_hours=overtime_period,
            overtime_period=overtime_period,
            security_status=security_info,
            suspicious_reason=suspicious_reason,
            work_content=work_content,
            is_holiday=is_holiday,
        )
