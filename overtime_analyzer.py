import sys
import pandas as pd
from datetime import datetime, time, timedelta
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


class OvertimeAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("초과근무 분석기")
        self.setGeometry(100, 100, 1200, 700)
        self.security_df = None  # 경비 기록 데이터프레임
        self.overtime_df = None  # 초과근무 기록 데이터프레임
        self.suspicious_records = []
        self.init_ui()

    def init_ui(self):
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
        overtime_file_group = QGroupBox("초과근무 기록 엑셀 파일 선택")
        overtime_file_layout = QHBoxLayout()

        self.overtime_file_label = QLabel("선택된 파일 없음")
        self.overtime_browse_button = QPushButton("파일 선택")
        self.overtime_browse_button.clicked.connect(lambda: self.browse_file("overtime"))

        overtime_file_layout.addWidget(self.overtime_file_label)
        overtime_file_layout.addWidget(self.overtime_browse_button)
        overtime_file_group.setLayout(overtime_file_layout)

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
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(
            ["날짜", "직원명", "초과근무 시간", "경비상태", "의심 사유"]
        )
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # 레이아웃에 위젯 추가
        main_layout.addWidget(security_file_group)
        main_layout.addWidget(overtime_file_group)
        main_layout.addWidget(date_group)
        main_layout.addWidget(self.analyze_button)
        main_layout.addWidget(self.table)
        main_layout.addWidget(self.export_button)

    def browse_file(self, file_type):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", "", "Excel Files (*.xlsx *.xls)", options=options
        )

        if file_path:
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
                    if file_ext == ".xls":
                        self.overtime_df = pd.read_excel(file_path, engine="xlrd")
                    else:
                        self.overtime_df = pd.read_excel(file_path, engine="openpyxl")
                    QMessageBox.information(
                        self,
                        "성공",
                        f"초과근무 기록 파일을 로드했습니다.\n총 {len(self.overtime_df)} 행의 데이터가 있습니다.",
                    )
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
        if self.security_df is None or self.overtime_df is None:
            QMessageBox.warning(self, "경고", "두 파일이 모두 로드되어야 합니다.")
            return

        try:
            # 경비 기록 파일 분석
            security_df_processed = self.process_security_log(self.security_df)

            # 초과근무 기록 파일 분석
            overtime_df_processed = self.process_overtime_log(self.overtime_df)

            # 두 데이터 비교 분석
            suspicious_records = self.compare_security_and_overtime(
                security_df_processed, overtime_df_processed
            )

            # 결과 테이블에 표시
            self.display_results(suspicious_records)

            # 분석 결과 메시지 표시
            if len(suspicious_records) == 0:
                QMessageBox.information(self, "분석 완료", "의심스러운 초과근무 기록이 없습니다.")
            else:
                QMessageBox.information(
                    self,
                    "분석 완료",
                    f"{len(suspicious_records)}개의 의심스러운 초과근무 기록을 발견했습니다.",
                )

            # 내보내기 버튼 활성화
            self.export_button.setEnabled(len(suspicious_records) > 0)

        except Exception as e:
            QMessageBox.critical(self, "오류", f"데이터 분석 중 오류가 발생했습니다: {str(e)}")
            import traceback

            print(traceback.format_exc())  # 상세 오류 정보 출력

    def process_security_log(self, df):
        """경비 기록을 처리하여 각 날짜별 경비 상태를 분석합니다."""
        try:
            # 필수 컬럼 정의
            required_cols = ["발생일자", "발생시각", "모드"]

            # 열 이름으로 컬럼 찾기
            col_mapping = {}

            # 컬럼 이름 매핑 (정확한 이름 또는 포함된 문자열로 찾기)
            for col in df.columns:
                col_str = str(col).lower()  # 컬럼명을 소문자로 변환하여 비교
                if "발생일자" in col_str or "날짜" in col_str:
                    col_mapping["발생일자"] = col
                elif "발생시각" in col_str or "시간" in col_str:
                    col_mapping["발생시각"] = col
                elif "모드" in col_str or "상태" in col_str or "내용" in col_str:
                    col_mapping["모드"] = col

            # 찾지 못한 컬럼은 기본 위치로 설정
            if "발생일자" not in col_mapping:
                col_mapping["발생일자"] = df.columns[0]  # A열
            if "발생시각" not in col_mapping:
                col_mapping["발생시각"] = df.columns[1]  # B열
            if "모드" not in col_mapping:
                if len(df.columns) > 8:
                    col_mapping["모드"] = df.columns[8]  # I열
                else:
                    col_mapping["모드"] = None

            # 필요한 컬럼이 없으면 오류 반환
            missing_cols = []
            if col_mapping["모드"] is None:
                missing_cols.append("모드")
            if missing_cols:
                raise ValueError(f"다음 컬럼을 찾을 수 없습니다: {', '.join(missing_cols)}")

            # 날짜 필터링 적용
            start_date = self.start_date.date().toString("yyyy-MM-dd")
            end_date = self.end_date.date().toString("yyyy-MM-dd")

            # 데이터프레임에서 날짜 열이 문자열이면 datetime으로 변환
            if not pd.api.types.is_datetime64_any_dtype(df[col_mapping["발생일자"]]):
                df[col_mapping["발생일자"]] = pd.to_datetime(
                    df[col_mapping["발생일자"]], errors="coerce"
                )

            # 시각 열이 문자열이면 datetime.time으로 변환
            if not pd.api.types.is_datetime64_any_dtype(df[col_mapping["발생시각"]]):
                try:
                    df["시간_datetime"] = pd.to_datetime(
                        df[col_mapping["발생시각"]], errors="coerce"
                    )
                    df["시간_시"] = df["시간_datetime"].dt.hour
                    df["시간_분"] = df["시간_datetime"].dt.minute
                except:
                    # 시간이 이미 시:분:초 형식인 경우
                    df["시간_시"] = df[col_mapping["발생시각"]].str.split(":").str[0].astype(int)
                    df["시간_분"] = df[col_mapping["발생시각"]].str.split(":").str[1].astype(int)
            else:
                df["시간_시"] = df[col_mapping["발생시각"]].dt.hour
                df["시간_분"] = df[col_mapping["발생시각"]].dt.minute

            # 필터링 적용
            filtered_df = df
            if start_date and end_date:
                filtered_df = df[
                    (df[col_mapping["발생일자"]] >= start_date)
                    & (df[col_mapping["발생일자"]] <= end_date)
                ]

            # 필요한 열만 선택하여 메모리 사용 최적화
            filtered_df_slim = filtered_df[
                [
                    col_mapping["발생일자"],
                    col_mapping["발생시각"],
                    col_mapping["모드"],
                    "시간_시",
                    "시간_분",
                ]
            ].copy()

            # 이전날짜와 다음날짜 관계 분석을 위해 날짜별로 정렬
            filtered_df_slim = filtered_df_slim.sort_values(
                by=[col_mapping["발생일자"], col_mapping["발생시각"]]
            )

            # 각 기록에 시간대 태그 생성 (새벽 4시 기준)
            filtered_df_slim["시간대"] = "주간"
            filtered_df_slim.loc[
                (filtered_df_slim["시간_시"] >= 0) & (filtered_df_slim["시간_시"] < 4), "시간대"
            ] = "새벽"

            # 업무일 계산 (새벽 4시 기준)
            filtered_df_slim["업무일"] = filtered_df_slim[col_mapping["발생일자"]].dt.date
            # 새벽 시간대(0-4시)는 전날의 업무일로 계산
            filtered_df_slim.loc[filtered_df_slim["시간대"] == "새벽", "업무일"] = (
                filtered_df_slim.loc[
                    filtered_df_slim["시간대"] == "새벽", col_mapping["발생일자"]
                ].dt.date
                - pd.Timedelta(days=1)
            )

            # 기록을 경비해제/경비시작으로 판단하는 함수
            def determine_record_type(row):
                mode = str(row[col_mapping["모드"]]).lower()

                # 명시적인 경비 해제/설정 상태 확인
                if "출근" in mode or "해제" in mode:
                    return "경비해제"
                elif "퇴근" in mode or "세팅" in mode or "세트" in mode:
                    return "경비시작"
                # 출입 기록은 컨텍스트로 판단해야 하므로 일단 불명확으로 분류
                elif "출입" in mode:
                    return "출입(불명확)"

                return "기타"

            # 각 기록 유형 판단
            filtered_df_slim["기록유형"] = filtered_df_slim.apply(determine_record_type, axis=1)

            # 출입(불명확) 기록 처리
            # 컨텍스트를 바탕으로 출입 기록을 경비해제 또는 경비시작으로 재분류
            business_days = filtered_df_slim["업무일"].unique()

            for business_day in business_days:
                day_records = filtered_df_slim[filtered_df_slim["업무일"] == business_day]
                unclear_records = day_records[day_records["기록유형"] == "출입(불명확)"]

                if len(unclear_records) > 0:
                    # 업무일 내 첫 기록과 마지막 기록이 불명확한 경우 처리
                    day_records_sorted = day_records.sort_values(
                        by=[col_mapping["발생일자"], col_mapping["발생시각"]]
                    )

                    # 첫 기록이 '출입(불명확)'이면 '경비해제'로 간주
                    if day_records_sorted.iloc[0]["기록유형"] == "출입(불명확)":
                        first_record_index = day_records_sorted.index[0]
                        filtered_df_slim.loc[first_record_index, "기록유형"] = "경비해제"

                    # 마지막 기록이 '출입(불명확)'이면 '경비시작'으로 간주
                    if day_records_sorted.iloc[-1]["기록유형"] == "출입(불명확)":
                        last_record_index = day_records_sorted.index[-1]
                        filtered_df_slim.loc[last_record_index, "기록유형"] = "경비시작"

            # 각 업무일별 경비 상태 시간 분석
            security_status_by_day = {}

            for business_day in business_days:
                day_records = filtered_df_slim[filtered_df_slim["업무일"] == business_day]
                day_records_sorted = day_records.sort_values(
                    by=[col_mapping["발생일자"], col_mapping["발생시각"]]
                )

                # 해당 업무일의 경비 상태 시간 기록 초기화
                security_status = []

                for _, record in day_records_sorted.iterrows():
                    record_date = record[col_mapping["발생일자"]]
                    record_time = f"{record['시간_시']:02d}:{record['시간_분']:02d}"
                    record_type = record["기록유형"]

                    # 경비해제/경비시작 시간 기록
                    if record_type == "경비해제":
                        security_status.append(
                            {
                                "시간": f"{record_date.strftime('%Y-%m-%d')} {record_time}",
                                "상태": "해제",
                            }
                        )
                    elif record_type == "경비시작":
                        security_status.append(
                            {
                                "시간": f"{record_date.strftime('%Y-%m-%d')} {record_time}",
                                "상태": "시작",
                            }
                        )

                # 업무일별 경비 상태 저장
                security_status_by_day[business_day] = security_status

            return security_status_by_day

        except Exception as e:
            # 예외 발생 시 상세 정보 출력하고 다시 발생
            import traceback

            print(f"경비 기록 처리 중 오류: {str(e)}")
            print(traceback.format_exc())
            raise

    def process_overtime_log(self, df):
        """초과근무 기록을 처리합니다."""
        try:
            # 열 이름으로 컬럼 찾기
            col_mapping = {}

            # 컬럼 이름 매핑 (정확한 이름 또는 포함된 문자열로 찾기)
            for col in df.columns:
                col_str = str(col).lower()  # 컬럼명을 소문자로 변환하여 비교
                if "날짜" in col_str or "일자" in col_str or "근무일" in col_str:
                    col_mapping["날짜"] = col
                elif "시작" in col_str and ("시간" in col_str or "시각" in col_str):
                    col_mapping["시작시간"] = col
                elif "종료" in col_str and ("시간" in col_str or "시각" in col_str):
                    col_mapping["종료시간"] = col
                elif "이름" in col_str or "성명" in col_str or "직원" in col_str:
                    col_mapping["이름"] = col

            # 필요한 컬럼이 없으면 자동 추정
            if "날짜" not in col_mapping:
                QMessageBox.warning(
                    self,
                    "경고",
                    "초과근무 날짜 컬럼을 찾을 수 없습니다. 첫 번째 컬럼을 사용합니다.",
                )
                col_mapping["날짜"] = df.columns[0]
            if "시작시간" not in col_mapping:
                QMessageBox.warning(
                    self,
                    "경고",
                    "초과근무 시작시간 컬럼을 찾을 수 없습니다. 두 번째 컬럼을 사용합니다.",
                )
                col_mapping["시작시간"] = df.columns[1] if len(df.columns) > 1 else None
            if "종료시간" not in col_mapping:
                QMessageBox.warning(
                    self,
                    "경고",
                    "초과근무 종료시간 컬럼을 찾을 수 없습니다. 세 번째 컬럼을 사용합니다.",
                )
                col_mapping["종료시간"] = df.columns[2] if len(df.columns) > 2 else None
            if "이름" not in col_mapping:
                QMessageBox.warning(self, "경고", "직원 이름 컬럼을 찾을 수 없습니다.")
                col_mapping["이름"] = None

            # 날짜 필터링 적용
            start_date = self.start_date.date().toString("yyyy-MM-dd")
            end_date = self.end_date.date().toString("yyyy-MM-dd")

            # 데이터프레임에서 날짜 열이 문자열이면 datetime으로 변환
            if not pd.api.types.is_datetime64_any_dtype(df[col_mapping["날짜"]]):
                df["날짜_datetime"] = pd.to_datetime(df[col_mapping["날짜"]], errors="coerce")
            else:
                df["날짜_datetime"] = df[col_mapping["날짜"]]

            # 필터링 적용
            filtered_df = df
            if start_date and end_date:
                filtered_df = df[
                    (df["날짜_datetime"] >= start_date) & (df["날짜_datetime"] <= end_date)
                ]

            # 초과근무 데이터 정리
            overtime_records = []

            # 정규 근무시간 설정 (9:00-18:00)
            regular_start = time(9, 0)
            regular_end = time(18, 0)

            for _, row in filtered_df.iterrows():
                try:
                    # 날짜 처리
                    work_date = row["날짜_datetime"].date()

                    # 시작 시간 처리
                    if pd.notna(row[col_mapping["시작시간"]]):
                        if isinstance(row[col_mapping["시작시간"]], str):
                            # 문자열 시간 처리 (형식: HH:MM 또는 HH:MM:SS)
                            start_parts = row[col_mapping["시작시간"]].split(":")
                            start_hour = int(start_parts[0])
                            start_minute = int(start_parts[1]) if len(start_parts) > 1 else 0
                            start_time = time(start_hour, start_minute)
                        elif isinstance(row[col_mapping["시작시간"]], datetime):
                            # datetime 객체인 경우
                            start_time = row[col_mapping["시작시간"]].time()
                        elif isinstance(row[col_mapping["시작시간"]], time):
                            # time 객체인 경우
                            start_time = row[col_mapping["시작시간"]]
                        else:
                            # 숫자인 경우 (예: 1830 -> 18:30)
                            time_value = int(row[col_mapping["시작시간"]])
                            start_hour = time_value // 100
                            start_minute = time_value % 100
                            start_time = time(start_hour, start_minute)
                    else:
                        # 시작 시간이 없으면 정규 근무 시작 시간으로 설정
                        start_time = regular_start

                    # 종료 시간 처리
                    if pd.notna(row[col_mapping["종료시간"]]):
                        if isinstance(row[col_mapping["종료시간"]], str):
                            # 문자열 시간 처리
                            end_parts = row[col_mapping["종료시간"]].split(":")
                            end_hour = int(end_parts[0])
                            end_minute = int(end_parts[1]) if len(end_parts) > 1 else 0
                            end_time = time(end_hour, end_minute)
                        elif isinstance(row[col_mapping["종료시간"]], datetime):
                            # datetime 객체인 경우
                            end_time = row[col_mapping["종료시간"]].time()
                        elif isinstance(row[col_mapping["종료시간"]], time):
                            # time 객체인 경우
                            end_time = row[col_mapping["종료시간"]]
                        else:
                            # 숫자인 경우 (예: 1830 -> 18:30)
                            time_value = int(row[col_mapping["종료시간"]])
                            end_hour = time_value // 100
                            end_minute = time_value % 100
                            end_time = time(end_hour, end_minute)
                    else:
                        # 종료 시간이 없으면 정규 근무 종료 시간으로 설정
                        end_time = regular_end

                    # 업무일 결정 (새벽 4시 기준)
                    business_date = work_date

                    # 자정 이후 새벽 4시 이전 근무는 전날 업무일로 계산
                    if end_time < time(4, 0):
                        business_date = work_date - timedelta(days=1)

                    # 초과근무 여부 확인 (시작 시간이 18시 이후이거나 종료 시간이 9시 이전)
                    is_overtime = start_time >= regular_end or end_time <= regular_start

                    # 업무 시간이 정규 근무시간(9-18)에 걸쳐있는 경우, 그 부분은 초과근무가 아님
                    if not is_overtime and (start_time < regular_end and end_time > regular_start):
                        # 시작 시간이 9시 이전이면 초과근무 시작 부분 기록
                        if start_time < regular_start:
                            overtime_start = start_time
                            overtime_end = regular_start

                            # 이름 정보
                            employee_name = (
                                row[col_mapping["이름"]]
                                if "이름" in col_mapping and col_mapping["이름"] is not None
                                else "Unknown"
                            )

                            overtime_records.append(
                                {
                                    "업무일": business_date,
                                    "날짜": work_date,
                                    "시작시간": overtime_start,
                                    "종료시간": overtime_end,
                                    "초과근무유형": "조기출근",
                                    "직원명": employee_name,
                                }
                            )

                        # 종료 시간이 18시 이후면 초과근무 종료 부분 기록
                        if end_time > regular_end:
                            overtime_start = regular_end
                            overtime_end = end_time

                            # 이름 정보
                            employee_name = (
                                row[col_mapping["이름"]]
                                if "이름" in col_mapping and col_mapping["이름"] is not None
                                else "Unknown"
                            )

                            overtime_records.append(
                                {
                                    "업무일": business_date,
                                    "날짜": work_date,
                                    "시작시간": overtime_start,
                                    "종료시간": overtime_end,
                                    "초과근무유형": "야근",
                                    "직원명": employee_name,
                                }
                            )
                    elif is_overtime:
                        # 전체가 초과근무인 경우
                        overtime_start = start_time
                        overtime_end = end_time

                        # 시간대에 따른 초과근무 유형 결정
                        if overtime_start < regular_start:
                            overtime_type = "조기출근"
                        else:
                            overtime_type = "야근"

                        # 이름 정보
                        employee_name = (
                            row[col_mapping["이름"]]
                            if "이름" in col_mapping and col_mapping["이름"] is not None
                            else "Unknown"
                        )

                        overtime_records.append(
                            {
                                "업무일": business_date,
                                "날짜": work_date,
                                "시작시간": overtime_start,
                                "종료시간": overtime_end,
                                "초과근무유형": overtime_type,
                                "직원명": employee_name,
                            }
                        )

                except Exception as e:
                    print(f"초과근무 기록 처리 중 오류 (무시됨): {str(e)}")
                    # 개별 레코드 처리 오류는 무시하고 계속 진행
                    continue

            return overtime_records

        except Exception as e:
            # 예외 발생 시 상세 정보 출력하고 다시 발생
            import traceback

            print(f"초과근무 기록 처리 중 오류: {str(e)}")
            print(traceback.format_exc())
            raise

    def compare_security_and_overtime(self, security_status_by_day, overtime_records):
        """경비 상태와 초과근무 기록을 비교 분석하여 의심스러운 기록을 찾습니다."""
        suspicious_records = []

        # 각 초과근무 기록에 대해 경비 상태 확인
        for overtime in overtime_records:
            business_date = overtime["업무일"]
            employee_name = overtime["직원명"]
            overtime_start = overtime["시작시간"]
            overtime_end = overtime["종료시간"]

            # 업무일에 해당하는 경비 기록 찾기
            security_status = security_status_by_day.get(business_date, [])

            if not security_status:
                # 해당 업무일의 경비 기록이 없으면 다음으로 넘어감
                continue

            # 초과근무 시간과 경비 상태 비교
            security_active = False  # 경비가 활성화된 상태인지
            suspicious_reason = None

            # 초과근무 시간을 datetime으로 변환
            overtime_start_dt = datetime.combine(overtime["날짜"], overtime_start)
            overtime_end_dt = datetime.combine(overtime["날짜"], overtime_end)

            # 자정을 넘어가는 경우 다음날로 설정
            if overtime_end < overtime_start:
                overtime_end_dt = datetime.combine(
                    overtime["날짜"] + timedelta(days=1), overtime_end
                )

            # 경비 기록을 시간순으로 정렬
            security_status.sort(key=lambda x: x["시간"])

            # 경비 상태 추적
            current_status = "시작"  # 기본값은 경비 활성화 상태(보수적 접근)

            # 초과근무 시작 시간 이전의 마지막 경비 상태 확인
            for record in security_status:
                record_time = datetime.strptime(record["시간"], "%Y-%m-%d %H:%M")

                if record_time < overtime_start_dt:
                    current_status = record["상태"]
                else:
                    break

            # 경비가 활성화된 상태에서 초과근무 기록이 있는지 확인
            if current_status == "시작":
                # 경비가 활성화된 동안의 초과근무 시간 계산
                overlap_start = overtime_start_dt
                overlap_end = overtime_end_dt

                # 초과근무 중 경비 상태가 변경되는지 확인
                for record in security_status:
                    record_time = datetime.strptime(record["시간"], "%Y-%m-%d %H:%M")

                    if overtime_start_dt <= record_time <= overtime_end_dt:
                        if record["상태"] == "해제":
                            # 경비가 해제되면 중첩 시간 종료
                            overlap_end = record_time
                            break

                # 경비 활성화 상태에서 초과근무 시간이 있으면 의심 기록
                if overlap_start < overlap_end:
                    overlap_duration = (
                        overlap_end - overlap_start
                    ).seconds / 3600  # 시간 단위로 변환

                    if overlap_duration >= 0.25:  # 15분(0.25시간) 이상 중첩되면 의심 기록으로 간주
                        suspicious_reason = (
                            f"경비 작동 중 {overlap_duration:.1f}시간 초과근무 기록 존재"
                        )
                        suspicious_records.append(
                            {
                                "날짜": business_date,
                                "직원명": employee_name,
                                "초과근무시간": f"{overtime_start.strftime('%H:%M')}-{overtime_end.strftime('%H:%M')}",
                                "경비상태": "경비 작동 중",
                                "의심사유": suspicious_reason,
                            }
                        )

        return suspicious_records

    def display_results(self, suspicious_records):
        """의심스러운 기록을 테이블에 표시합니다."""
        # 테이블 초기화
        self.table.setRowCount(0)
        self.suspicious_records = suspicious_records

        if not suspicious_records:
            return

        # 테이블에 행 추가
        for i, record in enumerate(suspicious_records):
            self.table.insertRow(i)

            # 날짜
            date_item = QTableWidgetItem(
                record["날짜"].strftime("%Y-%m-%d")
                if isinstance(record["날짜"], datetime.date)
                else str(record["날짜"])
            )
            self.table.setItem(i, 0, date_item)

            # 직원명
            name_item = QTableWidgetItem(str(record["직원명"]))
            self.table.setItem(i, 1, name_item)

            # 초과근무 시간
            time_item = QTableWidgetItem(str(record["초과근무시간"]))
            self.table.setItem(i, 2, time_item)

            # 경비 상태
            security_item = QTableWidgetItem(str(record["경비상태"]))
            self.table.setItem(i, 3, security_item)

            # 의심 사유
            reason_item = QTableWidgetItem(str(record["의심사유"]))
            self.table.setItem(i, 4, reason_item)

    def export_results(self):
        if not self.suspicious_records:
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
                    employee = self.table.item(row, 1).text()
                    overtime = self.table.item(row, 2).text()
                    security = self.table.item(row, 3).text()
                    reason = self.table.item(row, 4).text()

                    data.append(
                        {
                            "날짜": date,
                            "직원명": employee,
                            "초과근무 시간": overtime,
                            "경비상태": security,
                            "의심 사유": reason,
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
    window = OvertimeAnalyzer()
    window.show()
    sys.exit(app.exec_())
