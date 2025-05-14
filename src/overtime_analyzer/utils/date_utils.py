"""
날짜 및 시간 관련 유틸리티 함수를 제공합니다.
"""

import pandas as pd
from datetime import datetime, time, date, timedelta
from typing import Tuple, Optional, Union


def parse_time(
    time_value: Union[str, datetime, time, None], default_time: time = None
) -> Optional[time]:
    """
    다양한 형식의 시간값을 파싱하여 datetime.time 객체로 변환합니다.

    Args:
        time_value: 변환할 시간 값 (문자열, datetime, time 객체 중 하나)
        default_time: 변환 실패 시 반환할 기본 시간 값

    Returns:
        datetime.time: 파싱된 시간 객체 또는 default_time
    """
    if pd.isna(time_value):
        return default_time

    try:
        if isinstance(time_value, str):
            # 문자열 형식 시간 파싱
            time_str = time_value.strip()
            if ":" in time_str:
                parts = time_str.split(":")
                if len(parts) >= 2:
                    hour = int(parts[0])
                    minute = int(parts[1])
                    second = int(parts[2]) if len(parts) > 2 else 0
                    return time(hour, minute, second)
            return default_time
        elif isinstance(time_value, datetime):
            # datetime 객체에서 time 추출
            return time_value.time()
        elif isinstance(time_value, time):
            # 이미 time 객체인 경우
            return time_value
        else:
            # 지원하지 않는 형식
            print(f"지원하지 않는 시간 형식: {type(time_value)}")
            return default_time
    except Exception as e:
        print(f"시간 파싱 오류: {str(e)}")
        return default_time


def is_valid_time_format(time_str: str) -> bool:
    """
    시간 문자열이 유효한 형식(HH:mm)인지 검사합니다.

    Args:
        time_str: 검사할 시간 문자열

    Returns:
        bool: 유효성 여부
    """
    if not isinstance(time_str, str):
        return False
    try:
        # HH:mm 형식 검사
        parts = time_str.strip().split(":")
        if len(parts) != 2:
            return False
        hour, minute = int(parts[0]), int(parts[1])
        return 0 <= hour < 24 and 0 <= minute < 60
    except:
        return False


def calculate_business_date(record_date: date, record_hour: int) -> date:
    """
    새벽 4시 기준으로 업무일을 계산합니다.
    새벽 0시~4시는 전날의 업무일로 처리합니다.

    Args:
        record_date: 기록 날짜
        record_hour: 기록 시간의 시간 부분(hour)

    Returns:
        date: 계산된 업무일 날짜
    """
    # 0시 ~ 4시 사이는 전날 업무일로 계산
    if 0 <= record_hour < 4:
        return record_date - timedelta(days=1)
    return record_date
