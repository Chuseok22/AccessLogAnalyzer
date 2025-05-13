import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication,
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
import os


class AccessLogAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("출입 기록 분석기")
        self.setGeometry(100, 100, 900, 600)
        self.df = None
        self.suspicious_dates = []
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)

        # 파일 선택 영역
        file_group = QGroupBox("엑셀 파일 선택")
        file_layout = QHBoxLayout()

        self.file_label = QLabel("선택된 파일 없음")
        self.browse_button = QPushButton("파일 선택")
        self.browse_button.clicked.connect(self.browse_file)

        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.browse_button)
        file_group.setLayout(file_layout)

        # 날짜 필터 영역
        date_group = QGroupBox("날짜 범위 설정 (선택)")
        date_layout = QHBoxLayout()

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate().addMonths(-1))

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())

        date_layout.addWidget(QLabel("시작 날짜:"))
        date_layout.addWidget(self.start_date)
        date_layout.addWidget(QLabel("종료 날짜:"))
        date_layout.addWidget(self.end_date)
        date_group.setLayout(date_layout)

        # 분석 버튼
        self.analyze_button = QPushButton("분석 시작")
        self.analyze_button.clicked.connect(self.analyze_data)
        self.analyze_button.setEnabled(False)

        # 내보내기 버튼
        self.export_button = QPushButton("결과 내보내기")
        self.export_button.clicked.connect(self.export_results)
        self.export_button.setEnabled(False)

        # 결과 테이블
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["날짜", "출근 횟수", "퇴근 횟수", "상세 기록"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # 레이아웃에 위젯 추가
        main_layout.addWidget(file_group)
        main_layout.addWidget(date_group)
        main_layout.addWidget(self.analyze_button)
        main_layout.addWidget(self.table)
        main_layout.addWidget(self.export_button)

    def browse_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", "", "Excel Files (*.xlsx *.xls)", options=options
        )

        if file_path:
            self.file_label.setText(file_path)
            self.analyze_button.setEnabled(True)
            try:
                # XLS 파일 지원을 위해 엔진 자동 감지 활성화
                file_ext = os.path.splitext(file_path)[1].lower()
                if file_ext == ".xls":
                    # 구형 Excel 파일은 xlrd 엔진 사용
                    self.df = pd.read_excel(file_path, engine="xlrd")
                else:
                    # 신형 Excel 파일은 openpyxl 사용
                    self.df = pd.read_excel(file_path, engine="openpyxl")

                QMessageBox.information(
                    self,
                    "성공",
                    f"파일을 성공적으로 로드했습니다.\n총 {len(self.df)} 행의 데이터가 있습니다.",
                )
            except Exception as e:
                QMessageBox.critical(
                    self, "오류", f"파일을 로드하는 중 오류가 발생했습니다: {str(e)}"
                )
                self.file_label.setText("선택된 파일 없음")
                self.analyze_button.setEnabled(False)

    def analyze_data(self):
        if self.df is None:
            return

        try:
            # 필수 컬럼 정의
            required_cols = ["발생일자", "발생시각", "모드"]

            # 열 이름으로 컬럼 찾기
            col_mapping = {}

            # 컬럼 이름 매핑 (정확한 이름 또는 포함된 문자열로 찾기)
            for col in self.df.columns:
                col_str = str(col).lower()  # 컬럼명을 소문자로 변환하여 비교
                if "발생일자" in col_str or "날짜" in col_str:
                    col_mapping["발생일자"] = col
                elif "발생시각" in col_str or "시간" in col_str:
                    col_mapping["발생시각"] = col
                elif "모드" in col_str or "상태" in col_str or "내용" in col_str:
                    col_mapping["모드"] = col

            # 찾지 못한 컬럼은 기본 위치(A, B, I열)로 설정
            if "발생일자" not in col_mapping:
                col_mapping["발생일자"] = self.df.columns[0]  # A열
            if "발생시각" not in col_mapping:
                col_mapping["발생시각"] = self.df.columns[1]  # B열
            if "모드" not in col_mapping:
                if len(self.df.columns) > 8:
                    col_mapping["모드"] = self.df.columns[8]  # I열
                else:
                    col_mapping["모드"] = None

            # 필요한 컬럼이 없으면 사용자에게 알림
            missing_cols = []
            if col_mapping["모드"] is None:
                missing_cols.append("모드")
            if missing_cols:
                QMessageBox.warning(
                    self,
                    "컬럼 문제",
                    f"다음 컬럼을 찾을 수 없습니다: {', '.join(missing_cols)}",
                )
                return

            # 날짜 필터링 적용
            start_date = self.start_date.date().toString("yyyy-MM-dd")
            end_date = self.end_date.date().toString("yyyy-MM-dd")

            # 데이터프레임에서 날짜 열이 문자열이면 datetime으로 변환
            if not pd.api.types.is_datetime64_any_dtype(self.df[col_mapping["발생일자"]]):
                self.df[col_mapping["발생일자"]] = pd.to_datetime(
                    self.df[col_mapping["발생일자"]], errors="coerce"
                )

            filtered_df = self.df
            if start_date and end_date:
                filtered_df = self.df[
                    (self.df[col_mapping["발생일자"]] >= start_date)
                    & (self.df[col_mapping["발생일자"]] <= end_date)
                ]

            # 모드 분석 (세트/해제 또는 출근/퇴근)
            check_in_modes = ["해제", "출근"]
            check_out_modes = ["세트", "퇴근"]

            # 각 날짜별로 그룹핑하여 출입 기록 분석
            self.table.setRowCount(0)
            self.suspicious_dates = []

            # 날짜별로 그룹화
            # 필요한 열만 선택하여 메모리 사용 최적화
            filtered_df_slim = filtered_df[
                [col_mapping["발생일자"], col_mapping["발생시각"], col_mapping["모드"]]
            ]
            grouped = filtered_df_slim.groupby(pd.Grouper(key=col_mapping["발생일자"], freq="D"))

            row = 0
            for date, group in grouped:
                if len(group) == 0:
                    continue

                check_in_count = sum(group[col_mapping["모드"]].isin(check_in_modes))
                check_out_count = sum(group[col_mapping["모드"]].isin(check_out_modes))

                # 출/퇴근이 각각 1회씩 있으면 정상, 그 외에는 의심 날짜
                is_suspicious = check_in_count != 1 or check_out_count != 1

                if is_suspicious:
                    self.suspicious_dates.append(date.strftime("%Y-%m-%d"))

                    self.table.insertRow(row)

                    # 날짜 표시
                    date_item = QTableWidgetItem(date.strftime("%Y-%m-%d"))
                    self.table.setItem(row, 0, date_item)

                    # 출근 횟수
                    check_in_item = QTableWidgetItem(str(check_in_count))
                    self.table.setItem(row, 1, check_in_item)

                    # 퇴근 횟수
                    check_out_item = QTableWidgetItem(str(check_out_count))
                    self.table.setItem(row, 2, check_out_item)

                    # 상세 정보
                    details = []
                    for _, record in group.iterrows():
                        mode = record[col_mapping["모드"]]
                        time = record[col_mapping["발생시각"]]
                        if isinstance(time, pd.Timestamp):
                            time = time.strftime("%H:%M:%S")
                        details.append(f"{time} - {mode}")

                    details_item = QTableWidgetItem(", ".join(details))
                    self.table.setItem(row, 3, details_item)

                    row += 1

            self.export_button.setEnabled(True)

            if row == 0:
                QMessageBox.information(self, "분석 완료", "의심스러운 날짜가 없습니다.")
            else:
                QMessageBox.information(
                    self, "분석 완료", f"{row}개의 의심스러운 날짜를 발견했습니다."
                )

        except Exception as e:
            QMessageBox.critical(self, "오류", f"데이터 분석 중 오류가 발생했습니다: {str(e)}")

    def export_results(self):
        if not self.suspicious_dates:
            QMessageBox.information(self, "내보내기", "내보낼 결과가 없습니다.")
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self, "결과 저장", "", "Excel Files (*.xlsx)", options=options
        )

        if file_path:
            try:
                # 결과를 데이터프레임으로 변환
                data = []
                for row in range(self.table.rowCount()):
                    date = self.table.item(row, 0).text()
                    check_in_count = self.table.item(row, 1).text()
                    check_out_count = self.table.item(row, 2).text()
                    details = self.table.item(row, 3).text()

                    data.append(
                        {
                            "날짜": date,
                            "출근 횟수": check_in_count,
                            "퇴근 횟수": check_out_count,
                            "상세 기록": details,
                        }
                    )

                results_df = pd.DataFrame(data)
                results_df.to_excel(file_path, index=False)

                QMessageBox.information(
                    self, "완료", f"결과가 성공적으로 저장되었습니다:\n{file_path}"
                )

            except Exception as e:
                QMessageBox.critical(
                    self, "오류", f"결과 내보내기 중 오류가 발생했습니다: {str(e)}"
                )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AccessLogAnalyzer()
    window.show()
    sys.exit(app.exec_())
