"""
UI 모듈 - 초과근무 분석기의 사용자 인터페이스를 관리합니다.

이 모듈은 PyQt5를 사용하여 초과근무 분석기의 그래픽 사용자 인터페이스(GUI)를 구현합니다.
파일 선택, 날짜 범위 설정, 분석 실행 및 결과 표시와 같은 UI 관련 기능을 제공합니다.
"""

import os
import sys
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMessageBox,
    QHBoxLayout,
    QGroupBox,
    QDateEdit,
)
from PyQt5.QtCore import Qt, QDate

from ..services.analyzer_service import AnalyzerService
from ..services.security_processor import SecurityProcessor
from ..services.overtime_processor import OvertimeProcessor


class AnalyzerUI(QMainWindow):
    """
    초과근무 분석기의 메인 UI 클래스입니다.
    """

    def __init__(self):
        """UI 초기화 및 기본 설정"""
        super().__init__()
        self.setWindowTitle("초과근무 분석기")
        self.setGeometry(100, 100, 1200, 700)

        # 데이터 저장 변수 초기화
        self.security_df = None
        self.overtime_df = None
        self.suspicious_records = []

        # 서비스 객체 생성
        self.security_processor = SecurityProcessor()
        self.overtime_processor = OvertimeProcessor()
        self.analyzer_service = AnalyzerService()

        # UI 초기화
        self.init_ui()

    def init_ui(self):
        """UI 레이아웃 및 위젯 구성"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)

        # 경비 기록 파일 선택 영역
        security_file_group = QGroupBox("경비 기록 엑셀 파일 선택")
        security_file_layout = QHBoxLayout()

        self.security_file_label = QLabel("선택된 파일 없음")
        self.security_browse_button = QPushButton("파일 선택")
        self.security_browse_button.clicked.connect(lambda: self.browse_file("security"))

        security_file_layout.addWidget(self.security_file_label)
        security_file_layout.addWidget(self.security_browse_button)
        security_file_group.setLayout(security_file_layout)

        # 초과근무 기록 파일 선택 영역
        overtime_file_group = QGroupBox(
            "초과근무 기록 엑셀 파일 선택 (3행부터 데이터 시작, G열:일자, H열:출근, I열:퇴근)"
        )
        overtime_file_layout = QHBoxLayout()

        self.overtime_file_label = QLabel("선택된 파일 없음")
        self.overtime_browse_button = QPushButton("파일 선택")
        self.overtime_browse_button.clicked.connect(lambda: self.browse_file("overtime"))

        # 설명 레이블 추가
        overtime_info = QLabel(
            "※ 엑셀 파일은 3행부터 데이터가 시작되고, 초과근무일자(G열), 출근시간(H열), 퇴근시간(I열) 형식이어야 합니다."
        )
        overtime_info.setWordWrap(True)

        overtime_file_layout.addWidget(self.overtime_file_label)
        overtime_file_layout.addWidget(self.overtime_browse_button)
        overtime_file_layout.addWidget(overtime_info)
        overtime_file_group.setLayout(overtime_file_layout)

        # 날짜 필터 영역
        date_group = QGroupBox("날짜 범위 설정 (선택)")
        date_layout = QHBoxLayout()

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate().addYears(-1))

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())

        date_layout.addWidget(QLabel("시작 날짜:"))
        date_layout.addWidget(self.start_date)
        date_layout.addWidget(QLabel("종료 날짜:"))
        date_layout.addWidget(self.end_date)
        date_group.setLayout(date_layout)

        # 분석 버튼
        analyze_btn_group = QGroupBox("분석 실행")
        analyze_btn_layout = QHBoxLayout()

        self.analyze_button = QPushButton("분석 시작")
        self.analyze_button.clicked.connect(self.analyze_data)
        self.analyze_button.setEnabled(False)

        analyze_hint = QLabel("두 파일이 모두 로드되면 분석 버튼이 활성화됩니다.")

        analyze_btn_layout.addWidget(self.analyze_button)
        analyze_btn_layout.addWidget(analyze_hint)
        analyze_btn_group.setLayout(analyze_btn_layout)

        # 내보내기 버튼
        self.export_button = QPushButton("결과 내보내기")
        self.export_button.clicked.connect(self.export_results)
        self.export_button.setEnabled(False)

        # 결과 테이블
        self.table = QTableWidget()
        self.table.setColumnCount(
            8
        )  # 날짜, 직원명, 부서명, 초과근무시간, 경비상태, 의심사유, 근무내용, 휴일여부
        self.table.setHorizontalHeaderLabels(
            [
                "날짜",
                "직원명",
                "부서명",
                "초과근무 시간",
                "경비상태",
                "의심 사유",
                "근무내용",
                "휴일여부",
            ]
        )
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # 레이아웃에 위젯 추가
        main_layout.addWidget(security_file_group)
        main_layout.addWidget(overtime_file_group)
        main_layout.addWidget(date_group)
        main_layout.addWidget(analyze_btn_group)
        main_layout.addWidget(self.table)
        main_layout.addWidget(self.export_button)

    def browse_file(self, file_type):
        """
        엑셀 파일 선택 대화상자를 표시하고 선택된 파일을 로드합니다.

        Args:
            file_type (str): 파일 유형 ('security' 또는 'overtime')
        """
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", "", "Excel Files (*.xlsx *.xls)", options=options
        )

        if file_path:  # file_path 가 존재하는 경우
            if file_type == "security":
                self.security_file_label.setText(file_path)
                try:
                    file_ext = os.path.splitext(file_path)[1].lower()
                    if file_ext == ".xls":
                        self.security_df = pd.read_excel(file_path, engine="xlrd")
                    else:
                        self.security_df = pd.read_excel(file_path, engine="openpyxl")
                    QMessageBox.information(
                        self,
                        "성공",
                        f"경비 기록 파일을 로드했습니다.\n총 {len(self.security_df)} 행의 데이터가 있습니다.",
                    )
                except Exception as e:
                    QMessageBox.critical(
                        self, "오류", f"파일을 로드하는 중 오류가 발생했습니다: {str(e)}"
                    )
                    self.security_file_label.setText("선택된 파일 없음")
                    self.security_df = None

            elif file_type == "overtime":
                self.overtime_file_label.setText(file_path)
                try:
                    file_ext = os.path.splitext(file_path)[1].lower()

                    # 엑셀 파일 로드 (헤더 없음으로 처리)
                    if file_ext == ".xls":
                        self.overtime_df = pd.read_excel(file_path, engine="xlrd", header=None)
                    else:
                        self.overtime_df = pd.read_excel(file_path, engine="openpyxl", header=None)

                    # 데이터 유효성 확인
                    if len(self.overtime_df) < 3:
                        QMessageBox.warning(
                            self,
                            "경고",
                            "초과근무 기록 파일에 데이터가 충분하지 않습니다. 최소 3행 이상의 데이터가 필요합니다.",
                        )
                    else:
                        # 데이터 구조 표시
                        data_info = f"초과근무 기록 파일을 로드했습니다.\n"
                        data_info += (
                            f"총 {len(self.overtime_df)} 행 중 첫 2행은 헤더로 무시됩니다.\n"
                        )
                        data_info += f"실제 처리될 데이터는 {len(self.overtime_df) - 2}개 행입니다."

                        QMessageBox.information(self, "성공", data_info)
                except Exception as e:
                    QMessageBox.critical(
                        self, "오류", f"파일을 로드하는 중 오류가 발생했습니다: {str(e)}"
                    )
                    self.overtime_file_label.setText("선택된 파일 없음")
                    self.overtime_df = None

            # 두 파일이 모두 로드되었을 때만 분석 버튼 활성화
            if self.security_df is not None and self.overtime_df is not None:
                self.analyze_button.setEnabled(True)

    def analyze_data(self):
        """
        경비 및 초과 근무 데이터를 분석하여 의심스러운 기록을 감지합니다.
        """
        if self.security_df is None or self.overtime_df is None:
            QMessageBox.warning(self, "경고", "두 파일이 모두 로드되어야 합니다.")
            return

        try:
            # 날짜 범위 설정
            start_date = self.start_date.date().toString("yyyy-MM-dd")
            end_date = self.end_date.date().toString("yyyy-MM-dd")

            # 경비 기록 파일 분석
            print("[DEBUG] 경비 기록 파일 처리 시작...")
            security_records = self.security_processor.process_security_log(
                self.security_df, start_date, end_date
            )
            print("[DEBUG] 경비 기록 파일 처리 완료")

            # 초과근무 기록 파일 분석
            print("[DEBUG] 초과근무 기록 파일 처리 시작...")
            overtime_records = self.overtime_processor.process_overtime_log(
                self.overtime_df, start_date, end_date
            )
            print("[DEBUG] 초과근무 기록 파일 처리 완료")

            # 두 데이터 비교 분석
            print("[DEBUG] 데이터 비교 분석 시작...")
            self.suspicious_records = self.analyzer_service.compare_security_and_overtime(
                security_records, overtime_records
            )
            print("[DEBUG] 데이터 비교 분석 완료")

            # 결과 테이블에 표시
            self.display_results(self.suspicious_records)

            # 분석 결과 메시지 표시
            if len(self.suspicious_records) == 0:
                QMessageBox.information(self, "분석 완료", "의심스러운 초과근무 기록이 없습니다.")
            else:
                QMessageBox.information(
                    self,
                    "분석 완료",
                    f"{len(self.suspicious_records)}개의 의심스러운 초과근무 기록을 발견했습니다.",
                )

            # 내보내기 버튼 활성화
            self.export_button.setEnabled(len(self.suspicious_records) > 0)

        except Exception as e:
            QMessageBox.critical(self, "오류", f"데이터 분석 중 오류가 발생했습니다: {str(e)}")
            import traceback

            print(traceback.format_exc())  # 상세 오류 정보 출력

    def display_results(self, suspicious_records):
        """
        의심스러운 기록을 테이블에 표시합니다.

        Args:
            suspicious_records (list): 의심스러운 초과근무 기록 목록
        """
        # 테이블 초기화
        self.table.setRowCount(0)

        if not suspicious_records:
            return

        # 테이블에 행 추가
        for i, record in enumerate(suspicious_records):
            self.table.insertRow(i)

            # 날짜
            date_item = QTableWidgetItem(
                record.record_date.strftime("%Y-%m-%d")
                if hasattr(record.record_date, "strftime")
                else str(record.record_date)
            )
            self.table.setItem(i, 0, date_item)

            # 직원명
            name_item = QTableWidgetItem(str(record.employee_name))
            self.table.setItem(i, 1, name_item)

            # 부서명
            dept_item = QTableWidgetItem(str(record.department))
            self.table.setItem(i, 2, dept_item)

            # 초과근무 시간
            time_item = QTableWidgetItem(str(record.overtime_hours))
            self.table.setItem(i, 3, time_item)

            # 경비 상태
            security_item = QTableWidgetItem(str(record.security_status))
            self.table.setItem(i, 4, security_item)

            # 의심 사유
            reason_item = QTableWidgetItem(str(record.suspicious_reason))
            self.table.setItem(i, 5, reason_item)

            # 근무내용
            content_item = QTableWidgetItem(str(record.work_content))
            self.table.setItem(i, 6, content_item)

            # 휴일여부
            holiday_item = QTableWidgetItem("휴일" if record.is_holiday else "평일")
            self.table.setItem(i, 7, holiday_item)

    def export_results(self):
        """결과를 엑셀 파일로 내보냅니다."""
        # 의심 기록이 없는 경우 처리
        if not self.suspicious_records:
            QMessageBox.information(self, "내보내기", "내보낼 의심 기록이 없습니다.")
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self, "결과 저장", "", "Excel Files (*.xlsx)", options=options
        )

        if not file_path:
            return

        try:
            # 의심 기록만 내보내기
            self.export_suspicious_records(file_path)
            QMessageBox.information(self, "완료", f"결과가 성공적으로 저장되었습니다:\n{file_path}")

        except Exception as e:
            QMessageBox.critical(self, "오류", f"결과 내보내기 중 오류가 발생했습니다: {str(e)}")

    def export_suspicious_records(self, file_path):
        """
        의심 기록만 엑셀로 내보냅니다.

        Args:
            file_path (str): 저장할 엑셀 파일 경로
        """
        # 결과를 데이터프레임으로 변환
        data = []
        for row in range(self.table.rowCount()):
            date = self.table.item(row, 0).text()
            employee = self.table.item(row, 1).text()
            department = self.table.item(row, 2).text() if self.table.item(row, 2) else ""
            overtime = self.table.item(row, 3).text()
            security = self.table.item(row, 4).text()
            reason = self.table.item(row, 5).text()
            work_content = self.table.item(row, 6).text() if self.table.item(row, 6) else ""
            holiday_status = self.table.item(row, 7).text() if self.table.item(row, 7) else "평일"

            data.append(
                {
                    "날짜": date,
                    "직원명": employee,
                    "부서명": department,
                    "초과근무 시간": overtime,
                    "경비상태": security,
                    "의심 사유": reason,
                    "근무내용": work_content,
                    "휴일여부": holiday_status,
                }
            )

        # 엑셀로 저장
        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)
