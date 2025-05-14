#!/usr/bin/env python3
# 초과근무 분석 로직만 테스트하는 스크립트
import pandas as pd
import os
from datetime import datetime, timedelta, time


def test_holiday_detection():
    """휴일 여부 판단 로직을 테스트합니다"""
    print("===== 휴일 여부 판단 로직 테스트 =====")

    # 테스트 케이스
    test_cases = [
        {"휴일여부": "Y", "expected": True, "desc": "Y값"},
        {"휴일여부": "y", "expected": True, "desc": "소문자 y값"},
        {"휴일여부": "휴일", "expected": True, "desc": "휴일 텍스트"},
        {"휴일여부": "토요일", "expected": True, "desc": "토요일"},
        {"휴일여부": "일요일", "expected": True, "desc": "일요일"},
        {"휴일여부": "공휴일", "expected": True, "desc": "공휴일"},
        {"휴일여부": "N", "expected": False, "desc": "N값"},
        {"휴일여부": "평일", "expected": False, "desc": "평일 텍스트"},
        {"휴일여부": "", "expected": False, "desc": "빈 값"},
    ]

    # 테스트 실행
    for i, test in enumerate(test_cases):
        holiday_value = str(test["휴일여부"]).strip().lower()
        is_holiday = any(x in holiday_value for x in ["y", "휴", "공휴", "토요일", "일요일"])
        result = "성공" if is_holiday == test["expected"] else "실패"
        print(
            f"[{i+1}] {test['desc']}: '{test['휴일여부']}' => {is_holiday} (예상: {test['expected']}) - {result}"
        )


def test_suspicious_overtime_detection():
    """경비 설정 후 초과근무 의심 구간 감지 로직을 테스트합니다"""
    print("\n===== 경비 설정 후 초과근무 감지 테스트 =====")

    # 테스트 케이스 1: 3월 27일 사례 - 21:16에 경비 설정 후 23:00까지 초과근무
    test_date = datetime(2025, 3, 27).date()
    security_set_time = datetime.combine(test_date, time(21, 16, 9))
    overtime_start = datetime.combine(test_date, time(18, 0))
    overtime_end = datetime.combine(test_date, time(23, 0))

    # 의심 구간 계산
    suspicious_start = max(security_set_time, overtime_start)
    suspicious_end = overtime_end
    suspicious_duration = (suspicious_end - suspicious_start).seconds / 3600

    print(f"날짜: {test_date}")
    print(f"경비 설정: {security_set_time.strftime('%H:%M:%S')}")
    print(f"초과근무: {overtime_start.strftime('%H:%M')} - {overtime_end.strftime('%H:%M')}")
    print(
        f"의심 구간: {suspicious_start.strftime('%H:%M:%S')} - {suspicious_end.strftime('%H:%M')} ({suspicious_duration:.2f}시간)"
    )
    print(
        f"의심 사유: 경비 작동 중 총 {suspicious_duration:.1f}시간 초과근무 기록 존재 ({suspicious_start.strftime('%H:%M')}-{suspicious_end.strftime('%H:%M')})"
    )

    # 테스트 케이스 2: 경비 설정이 여러번인 경우
    print("\n----- 여러 경비 설정 시간이 있는 케이스 -----")
    test_date = datetime(2025, 5, 1).date()
    security_changes = [
        {"시간": datetime.combine(test_date, time(18, 30)), "상태": "해제"},  # 18:30 경비 해제
        {"시간": datetime.combine(test_date, time(19, 45)), "상태": "시작"},  # 19:45 경비 설정
        {"시간": datetime.combine(test_date, time(20, 15)), "상태": "해제"},  # 20:15 경비 해제
        {"시간": datetime.combine(test_date, time(21, 30)), "상태": "시작"},  # 21:30 경비 설정
    ]
    overtime_start = datetime.combine(test_date, time(18, 0))
    overtime_end = datetime.combine(test_date, time(22, 30))

    # 의심 구간 계산
    suspicious_intervals = []

    # CASE 1: 경비 설정(시작) 후 초과근무가 계속된 경우 감지
    for i, change in enumerate(security_changes):
        if change["상태"] == "시작":  # 경비 설정(시작)
            security_set_time = change["시간"]
            # 이 시각 이후의 초과근무는 의심스러움
            if security_set_time < overtime_end:
                start_time = max(security_set_time, overtime_start)
                end_time = overtime_end
                suspicious_intervals.append((start_time, end_time))
                print(
                    f"경비설정({security_set_time.strftime('%H:%M:%S')}) 이후 초과근무 감지: {start_time.strftime('%H:%M:%S')}-{end_time.strftime('%H:%M:%S')}"
                )

    # 모든 의심 구간에 대해 총 중첩 시간 계산
    total_suspicious_hours = 0
    suspicious_periods = []

    for start_time, end_time in suspicious_intervals:
        if start_time < end_time:  # 유효한 구간만 처리
            duration_hours = (end_time - start_time).seconds / 3600
            if duration_hours >= 0.25:  # 15분(0.25시간) 이상 중첩만 고려
                total_suspicious_hours += duration_hours
                start_str = start_time.strftime("%H:%M")
                end_str = end_time.strftime("%H:%M")
                suspicious_periods.append(f"{start_str}-{end_str}")

    # 의심 시간 출력
    if suspicious_periods:
        period_str = ", ".join(suspicious_periods)
        print(f"총 의심 시간: {total_suspicious_hours:.2f}시간")
        print(f"의심 구간: {period_str}")
        print(
            f"의심 사유: 경비 작동 중 총 {total_suspicious_hours:.1f}시간 초과근무 기록 존재 ({period_str})"
        )


if __name__ == "__main__":
    test_holiday_detection()
    test_suspicious_overtime_detection()
