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
        self.table.setHorizontalHeaderLabels(
            ["날짜", "경비해제 횟수", "경비시작 횟수", "상세 기록"]
        )
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

            # 시각 열이 문자열이면 datetime.time으로 변환
            if not pd.api.types.is_datetime64_any_dtype(self.df[col_mapping["발생시각"]]):
                try:
                    self.df["시간_datetime"] = pd.to_datetime(
                        self.df[col_mapping["발생시각"]], errors="coerce"
                    )
                    self.df["시간_시"] = self.df["시간_datetime"].dt.hour
                except:
                    # 시간이 이미 시:분:초 형식인 경우
                    self.df["시간_시"] = (
                        self.df[col_mapping["발생시각"]].str.split(":").str[0].astype(int)
                    )
            else:
                self.df["시간_시"] = self.df[col_mapping["발생시각"]].dt.hour

            filtered_df = self.df
            if start_date and end_date:
                filtered_df = self.df[
                    (self.df[col_mapping["발생일자"]] >= start_date)
                    & (self.df[col_mapping["발생일자"]] <= end_date)
                ]

            # 필요한 열만 선택하여 메모리 사용 최적화
            filtered_df_slim = filtered_df[
                [col_mapping["발생일자"], col_mapping["발생시각"], col_mapping["모드"], "시간_시"]
            ].copy()

            # 이전날짜와 다음날짜 관계 분석을 위해 날짜별로 정렬
            filtered_df_slim = filtered_df_slim.sort_values(
                by=[col_mapping["발생일자"], col_mapping["발생시각"]]
            )

            # 각 기록에 오전/오후 태그 생성
            filtered_df_slim["시간대"] = "오후"
            filtered_df_slim.loc[filtered_df_slim["시간_시"] < 12, "시간대"] = "오전"
            filtered_df_slim.loc[filtered_df_slim["시간_시"] < 5, "시간대"] = "새벽"

            # 각 날짜별로 첫 이벤트와 마지막 이벤트 식별을 위한 준비
            filtered_df_slim["이벤트순서"] = filtered_df_slim.groupby(
                col_mapping["발생일자"]
            ).cumcount()

            # 각 기록을 경비해제/경비시작으로 판단하는 함수
            def determine_record_type(row):
                mode = str(row[col_mapping["모드"]]).lower()
                time_period = row["시간대"]
                hour = row["시간_시"]

                # 명시적인 경비 해제/설정 상태 확인
                if "출근" in mode or "해제" in mode:
                    return "경비해제"
                elif "퇴근" in mode or "세팅" in mode or "세트" in mode:
                    return "경비시작"
                # 출입 기록은 시간대에 따라 판단
                elif "출입" in mode:
                    # 출입은 시간대로만 판단이 어려움
                    # 로그를 기타로 분류하고 나중에 컨텍스트로 판단
                    return "출입(불명확)"

                return "기타"

            # 각 기록 유형 판단
            filtered_df_slim["기록유형"] = filtered_df_slim.apply(determine_record_type, axis=1)

            # '출입(불명확)' 기록을 같은 날짜의 다른 기록으로 분류 시도
            # 각 날짜별 첫번째와 마지막 '출입'은 시간에 따라 분류
            days = filtered_df_slim[col_mapping["발생일자"]].dt.date.unique()

            for day in days:
                day_records = filtered_df_slim[
                    filtered_df_slim[col_mapping["발생일자"]].dt.date == day
                ]
                unclear_records = day_records[day_records["기록유형"] == "출입(불명확)"]

                if len(unclear_records) > 0:
                    # 해당 날짜의 첫 기록과 마지막 기록이 '출입(불명확)'인 경우 시간대로 추정
                    first_event_idx = day_records["이벤트순서"].min()
                    last_event_idx = day_records["이벤트순서"].max()

                    # 첫 기록이 '출입(불명확)'이고 오전이면 '경비해제'로 간주
                    first_event = day_records[day_records["이벤트순서"] == first_event_idx].iloc[0]
                    if first_event["기록유형"] == "출입(불명확)" and first_event["시간대"] in [
                        "오전",
                        "새벽",
                    ]:
                        filtered_df_slim.loc[
                            (filtered_df_slim[col_mapping["발생일자"]].dt.date == day)
                            & (filtered_df_slim["이벤트순서"] == first_event_idx),
                            "기록유형",
                        ] = "경비해제"

                    # 마지막 기록이 '출입(불명확)'이고 오후이면 '경비시작'으로 간주
                    last_event = day_records[day_records["이벤트순서"] == last_event_idx].iloc[0]
                    if last_event["기록유형"] == "출입(불명확)" and last_event["시간대"] == "오후":
                        filtered_df_slim.loc[
                            (filtered_df_slim[col_mapping["발생일자"]].dt.date == day)
                            & (filtered_df_slim["이벤트순서"] == last_event_idx),
                            "기록유형",
                        ] = "경비시작"

            # 날짜별로 그룹화
            grouped = filtered_df_slim.groupby(pd.Grouper(key=col_mapping["발생일자"], freq="D"))

            # 분석 결과 저장
            analysis_results = []
            prev_date = None
            prev_date_data = None

            # 첫번째 패스: 날짜별 기본 분석
            for date, group in grouped:
                if len(group) == 0:
                    continue

                # 현재 날짜의 경비 해제/시작 기록 분석
                security_off_records = group[group["기록유형"] == "경비해제"]
                security_on_records = group[group["기록유형"] == "경비시작"]
                unclear_records = group[group["기록유형"] == "출입(불명확)"]

                security_off_count = len(security_off_records)
                security_on_count = len(security_on_records)
                unclear_count = len(unclear_records)

                # 결과 저장
                analysis_results.append(
                    {
                        "date": date,
                        "security_off_count": security_off_count,
                        "security_on_count": security_on_count,
                        "unclear_count": unclear_count,
                        "records": group,
                        "is_suspicious": False,  # 일단 의심 아님으로 초기화
                    }
                )

            # 두번째 패스: 자정 넘어가는 경우 고려
            for i, current in enumerate(analysis_results):
                # 정상 패턴: 경비해제 1회, 경비시작 1회
                # 의심 패턴: 경비해제나 경비시작이 0회 또는 2회 이상

                # 의심스러운 패턴 판단
                if current["security_off_count"] != 1 or current["security_on_count"] != 1:
                    # 경비해제는 있으나 경비시작이 없는 경우: 다음날 새벽 경비시작 확인
                    if current["security_off_count"] == 1 and current["security_on_count"] == 0:
                        # 다음날 데이터가 있는지 확인
                        if i + 1 < len(analysis_results):
                            next_day = analysis_results[i + 1]
                            # 다음날 새벽(5시 이전)에 경비시작 기록이 있는지 확인
                            next_day_early_records = next_day["records"][
                                (next_day["records"]["시간대"] == "새벽")
                                & (next_day["records"]["기록유형"] == "경비시작")
                            ]

                            # 다음날 새벽 경비시작이 있으면 정상으로 간주
                            if len(next_day_early_records) > 0:
                                current["is_suspicious"] = False
                                continue

                    # 불명확한 출입 기록이 있는 경우 추가 검증
                    elif current["unclear_count"] > 0:
                        # 불명확한 출입 기록만 있고 다른 기록이 없으면 의심 대상
                        if current["security_off_count"] == 0 and current["security_on_count"] == 0:
                            current["is_suspicious"] = True
                        # 경비해제/시작이 각각 1회씩 있고, 그 외에 불명확한 출입이 있으면 의심 대상
                        elif (
                            current["security_off_count"] == 1 and current["security_on_count"] == 1
                        ):
                            current["is_suspicious"] = True

                    # 그 외의 비정상 패턴은 의심스러운 날짜로 표시
                    if current["security_off_count"] > 1 or current["security_on_count"] > 1:
                        current["is_suspicious"] = True
                    # 경비해제/시작이 둘 다 없는 경우도 의심 대상
                    elif current["security_off_count"] == 0 and current["security_on_count"] == 0:
                        current["is_suspicious"] = True

            # 테이블 초기화
            self.table.setRowCount(0)
            self.suspicious_dates = []

            # 의심스러운 날짜만 테이블에 표시
            row = 0
            for result in analysis_results:
                if result["is_suspicious"]:
                    date = result["date"]
                    self.suspicious_dates.append(date.strftime("%Y-%m-%d"))

                    # 테이블에 행 추가
                    self.table.insertRow(row)

                    # 날짜 표시
                    date_item = QTableWidgetItem(date.strftime("%Y-%m-%d"))
                    self.table.setItem(row, 0, date_item)

                    # 경비해제 횟수
                    security_off_item = QTableWidgetItem(str(result["security_off_count"]))
                    self.table.setItem(row, 1, security_off_item)

                    # 경비시작 횟수
                    security_on_item = QTableWidgetItem(str(result["security_on_count"]))
                    self.table.setItem(row, 2, security_on_item)

                    # 상세 정보
                    details = []
                    for _, record in result["records"].iterrows():
                        mode = record[col_mapping["모드"]]
                        time = record[col_mapping["발생시각"]]
                        record_type = record["기록유형"]
                        if isinstance(time, pd.Timestamp):
                            time = time.strftime("%H:%M:%S")
                        details.append(f"{time} - {mode} ({record_type})")

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
            import traceback

            print(traceback.format_exc())  # 상세 오류 정보 출력

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
                            "경비해제 횟수": check_in_count,
                            "경비시작 횟수": check_out_count,
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
