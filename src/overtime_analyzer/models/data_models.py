"""
데이터 모델 모듈 - 분석에 필요한 데이터 구조를 정의합니다.

이 모듈은 초과근무 분석에 필요한 데이터 클래스와 구조를 제공합니다.
"""

from dataclasses import dataclass
from datetime import date, time, datetime
from typing import Dict, List, Optional, Any, Union


@dataclass
class SecurityRecord:
    """경비 기록 데이터 클래스"""

    date: date  # 발생일자
    time: time  # 발생시각
    mode: str  # 모드 (출근/퇴근/해제/세팅 등)
    business_date: date  # 업무일자 (새벽 4시 기준)
    record_type: str  # 기록 유형 (경비해제/경비시작/출입(불명확)/기타)


@dataclass
class SecurityStatusChange:
    """경비 상태 변화 기록 데이터 클래스"""

    time: datetime  # 변화 시간
    status: str  # 상태 (해제/시작)


@dataclass
class SecurityPeriod:
    """경비 상태 구간 데이터 클래스"""

    start: datetime  # 시작 시간
    end: datetime  # 종료 시간
    status: str  # 상태 (해제/시작)


@dataclass
class OvertimeRecord:
    """초과근무 기록 데이터 클래스"""

    business_date: date  # 업무일자 (새벽 4시 기준)
    original_date: date  # 원래 날짜
    start_time: time  # 시작 시간
    end_time: time  # 종료 시간
    overtime_type: str  # 초과근무 유형 (조기출근/야간근무/휴일근무)
    employee_name: str  # 직원 이름
    department: str  # 부서명
    recorded_overtime: float  # 기록된 초과근무시간
    work_content: str  # 근무 내용
    is_holiday: bool  # 휴일 여부


@dataclass
class SuspiciousRecord:
    """의심스러운 초과근무 기록 데이터 클래스"""

    date: date  # 날짜
    employee_name: str  # 직원명
    department: str  # 부서명
    overtime_hours: str  # 초과근무 시간 표시
    overtime_period: str  # 초과근무 시간 (시작-종료)
    security_status: str  # 경비 상태
    suspicious_reason: str  # 의심 사유
    work_content: str  # 근무 내용
    is_holiday: bool  # 휴일 여부
