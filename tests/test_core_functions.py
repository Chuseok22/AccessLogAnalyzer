"""
출입기록분석기 단위 테스트

이 파일은 리팩토링된 출입기록분석기의 핵심 기능에 대한 테스트를 수행합니다.
"""

import unittest
from datetime import datetime, time, date, timedelta
import pandas as pd

from src.overtime_analyzer.models.data_models import (
    SecurityRecord,
    OvertimeRecord,
    SuspiciousRecord,
)
from src.overtime_analyzer.services.security_processor import SecurityProcessor
from src.overtime_analyzer.services.overtime_processor import OvertimeProcessor
from src.overtime_analyzer.services.analyzer_service import AnalyzerService
from src.overtime_analyzer.utils.date_utils import is_business_day, parse_time


class TestSecurityProcessor(unittest.TestCase):
    """경비 처리 모듈 테스트"""

    def setUp(self):
        self.processor = SecurityProcessor()

        # 테스트용 데이터 생성
        self.test_data = pd.DataFrame(
            {
                "발생일자": ["2023-05-01", "2023-05-01", "2023-05-01", "2023-05-02"],
                "발생시각": ["08:30:00", "18:00:00", "22:00:00", "09:00:00"],
                "모드": ["경비해제", "경비설정", "경비해제", "경비해제"],
            }
        )

    def test_identify_columns(self):
        """컬럼 인식 기능을 테스트합니다."""
        col_mapping = self.processor._identify_columns(self.test_data)
        self.assertEqual(col_mapping["발생일자"], "발생일자")
        self.assertEqual(col_mapping["발생시각"], "발생시각")
        self.assertEqual(col_mapping["모드"], "모드")

    def test_parse_security_record(self):
        """보안 기록 파싱 기능을 테스트합니다."""
        test_row = self.test_data.iloc[0]
        record = self.processor._parse_security_record(
            test_row, {"발생일자": "발생일자", "발생시각": "발생시각", "모드": "모드"}
        )

        self.assertEqual(record.date, date(2023, 5, 1))
        self.assertEqual(record.time, time(8, 30, 0))
        self.assertEqual(record.mode, "경비해제")
        self.assertEqual(record.record_type, "해제")

    def test_determine_business_date(self):
        """업무일 계산 기능을 테스트합니다."""
        # 자정 이전 시간 (같은 날짜로 취급)
        midnight_before = datetime.combine(date(2023, 5, 1), time(23, 0))
        bd_before = self.processor._determine_business_date(midnight_before)
        self.assertEqual(bd_before, date(2023, 5, 1))

        # 자정 이후 새벽 시간 (전날 업무일로 취급)
        midnight_after = datetime.combine(date(2023, 5, 2), time(3, 0))
        bd_after = self.processor._determine_business_date(midnight_after)
        self.assertEqual(bd_after, date(2023, 5, 1))

        # 오전 4시 이후는 당일 업무일로 취급
        morning = datetime.combine(date(2023, 5, 2), time(5, 0))
        bd_morning = self.processor._determine_business_date(morning)
        self.assertEqual(bd_morning, date(2023, 5, 2))


class TestOvertimeProcessor(unittest.TestCase):
    """초과근무 처리 모듈 테스트"""

    def setUp(self):
        self.processor = OvertimeProcessor()

        # 테스트용 데이터 생성 (3행부터 데이터 시작)
        self.test_data = pd.DataFrame(
            {
                0: ["헤더1", "헤더2", "김사원", "이대리"],  # 3행부터 데이터
                1: ["정보1", "정보2", "IT부", "개발부"],
                6: ["2023-05-01", "2023-05-02", "2023-05-03", "2023-05-04"],  # G열: 근무일자
                7: ["18:00", "19:00", "18:30", "19:00"],  # H열: 시작시간
                8: ["22:00", "23:00", "23:30", "22:00"],  # I열: 종료시간
                10: ["개발 작업", "버그 수정", "테스트", "코드 리뷰"],  # K열: 근무내용
            }
        )

    def test_extract_overtime_data(self):
        """초과근무 데이터 추출 기능을 테스트합니다."""
        result = self.processor._extract_overtime_data(self.test_data)

        # 첫 번째 데이터 검증
        self.assertEqual(len(result), 2)  # 두 개의 데이터행만 추출
        self.assertEqual(result[0]["직원명"], "김사원")
        self.assertEqual(result[0]["부서명"], "IT부")
        self.assertEqual(result[0]["근무일자"].strftime("%Y-%m-%d"), "2023-05-03")
        self.assertEqual(result[0]["시작시간"].strftime("%H:%M"), "18:30")
        self.assertEqual(result[0]["종료시간"].strftime("%H:%M"), "23:30")
        self.assertEqual(result[0]["근무내용"], "테스트")


class TestAnalyzerService(unittest.TestCase):
    """분석 서비스 모듈 테스트"""

    def setUp(self):
        self.analyzer = AnalyzerService()

        # 테스트용 경비 상태 데이터
        self.security_status_by_day = {
            date(2023, 5, 1): [
                {"시간": "2023-05-01 08:00", "상태": "해제"},
                {"시간": "2023-05-01 18:00", "상태": "시작"},
                {"시간": "2023-05-01 22:00", "상태": "해제"},
            ],
            date(2023, 5, 2): [
                {"시간": "2023-05-02 08:00", "상태": "해제"},
                {"시간": "2023-05-02 18:00", "상태": "시작"},
            ],
        }

        # 테스트용 초과근무 기록
        self.overtime_records = [
            {
                "업무일": date(2023, 5, 1),
                "직원명": "김사원",
                "부서명": "IT부",
                "시작시간": time(18, 30),
                "종료시간": time(21, 30),
                "근무내용": "개발 작업",
                "휴일여부": False,
            },
            {
                "업무일": date(2023, 5, 2),
                "직원명": "이대리",
                "부서명": "개발부",
                "시작시간": time(19, 0),
                "종료시간": time(23, 0),
                "근무내용": "버그 수정",
                "휴일여부": False,
            },
        ]

    def test_compare_security_and_overtime(self):
        """경비상태와 초과근무 비교 기능을 테스트합니다."""
        results = self.analyzer.compare_security_and_overtime(
            self.security_status_by_day, self.overtime_records
        )

        # 검증
        self.assertEqual(len(results), 2)  # 두 개의 의심 기록이 나와야 함

        # 첫 번째 의심 기록 검증
        self.assertEqual(results[0].date, date(2023, 5, 1))
        self.assertEqual(results[0].employee_name, "김사원")
        self.assertEqual(results[0].department, "IT부")
        self.assertTrue("경비 작동 중" in results[0].security_status)

        # 두 번째 의심 기록 검증
        self.assertEqual(results[1].date, date(2023, 5, 2))
        self.assertEqual(results[1].employee_name, "이대리")
        self.assertEqual(results[1].department, "개발부")
        self.assertTrue("경비 작동 중" in results[1].security_status)


class TestDateUtils(unittest.TestCase):
    """날짜 유틸리티 테스트"""

    def test_is_business_day(self):
        """업무일 판단 기능을 테스트합니다."""
        # 평일 (화요일)
        tuesday = date(2023, 5, 2)
        self.assertTrue(is_business_day(tuesday))

        # 주말 (일요일)
        sunday = date(2023, 5, 7)
        self.assertFalse(is_business_day(sunday))

    def test_parse_time(self):
        """시간 파싱 기능을 테스트합니다."""
        # 정상 시간 형식
        time_obj = parse_time("18:30")
        self.assertEqual(time_obj, time(18, 30))

        # 초 포함 시간 형식
        time_with_seconds = parse_time("18:30:45")
        self.assertEqual(time_with_seconds, time(18, 30, 45))

        # 빈 문자열 예외처리
        empty_time = parse_time("")
        self.assertIsNone(empty_time)


if __name__ == "__main__":
    unittest.main()
