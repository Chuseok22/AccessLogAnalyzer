import sys
import pandas as pd
from datetime import datetime, time, timedelta, date
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
    QInputDialog,
)
from PyQt5.QtCore import Qt, QDate
import os


class OvertimeAnalyzer(QMainWindow):

    # OvertimeAnalyzer 객체 초기화 및 GUI 창의 기본 설정
    def __init__(self):
        super().__init__()
        self.setWindowTitle("초과근무 분석기")
        self.setGeometry(100, 100, 1200, 700)
        self.security_df = None  # 경비 기록 데이터프레임
        self.overtime_df = None  # 초과근무 기록 데이터프레임
        self.suspicious_records = []
        self.init_ui()

    # GUI의 레이아웃과 위젯 구성
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

    # 사용자 엑셀 파일 선택 및 로드
    def browse_file(self, file_type):
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

    # 경비 및 초과 근무 데이터를 분석하여 의심스러운 기록 탐지 (메인 메서드)
    def analyze_data(self):
        if self.security_df is None or self.overtime_df is None:
            QMessageBox.warning(self, "경고", "두 파일이 모두 로드되어야 합니다.")
            return

        try:
            # 경비 기록 파일 분석
            print("[DEBUG] 경비 기록 파일 처리 시작...")
            security_df_processed = self.process_security_log(self.security_df)
            print("[DEBUG] 경비 기록 파일 처리 완료")

            # 초과근무 기록 파일 분석
            print("[DEBUG] 초과근무 기록 파일 처리 시작...")
            overtime_df_processed = self.process_overtime_log(self.overtime_df)
            print("[DEBUG] 초과근무 기록 파일 처리 완료")

            # 두 데이터 비교 분석
            print("[DEBUG] 데이터 비교 분석 시작...")
            suspicious_records = self.compare_security_and_overtime(
                security_df_processed, overtime_df_processed
            )
            print("[DEBUG] 데이터 비교 분석 완료")

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

    # 경비 로그를 처리하여 날짜별 경비 상태 (설정/해제) 분석
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

            # 경비 데이터가 불명확한 업무일 기록
            unclear_security_days = []

            for business_day in business_days:
                day_records = filtered_df_slim[filtered_df_slim["업무일"] == business_day]
                unclear_records = day_records[day_records["기록유형"] == "출입(불명확)"]

                # 업무일 내 기록 시간순 정렬
                day_records_sorted = day_records.sort_values(
                    by=[col_mapping["발생일자"], col_mapping["발생시각"]]
                )

                # 경비 해제/시작 기록 확인
                has_release = any(
                    record["기록유형"] == "경비해제" for _, record in day_records_sorted.iterrows()
                )
                has_start = any(
                    record["기록유형"] == "경비시작" for _, record in day_records_sorted.iterrows()
                )

                if len(unclear_records) > 0:
                    # 새로운 요구사항에 맞게 처리:
                    # 1. 첫 기록이 '출입(불명확)'이면 '경비해제'로 간주
                    # 2. 그 외 모든 '출입(불명확)' 기록은 무시

                    # 첫 기록이 '출입(불명확)'이면 '경비해제'로 간주
                    if day_records_sorted.iloc[0]["기록유형"] == "출입(불명확)":
                        first_record_index = day_records_sorted.index[0]
                        filtered_df_slim.loc[first_record_index, "기록유형"] = "경비해제"
                        print(
                            f"[경비판단] {business_day} - 첫 기록이 '출입'이므로 '경비해제'로 판단"
                        )

                    # 첫 번째가 아닌 모든 '출입(불명확)' 기록은 무시 (기타로 변경)
                    for i, record in enumerate(day_records_sorted.iterrows()):
                        if i == 0:  # 첫 번째 기록은 건너뛰기 (이미 처리됨)
                            continue

                        idx, row = record
                        # 첫 번째가 아닌 모든 불명확 출입 기록은 '기타'로 처리 (무시)
                        if row["기록유형"] == "출입(불명확)":
                            filtered_df_slim.loc[idx, "기록유형"] = "기타"
                            print(
                                f"[출입무시] {business_day} - 첫 번째가 아닌 출입기록은 무시함 (시간: {row['시간_시']:02d}:{row['시간_분']:02d})"
                            )

                    # '출입(불명확)' 기록은 첫번째를 제외한 모든 기록이 '기타'로 처리되므로 검사 불필요

                    # 마지막 기록이 확실한 경비시작이 아니면 확인 필요
                    # (단, '기타'로 처리된 출입 기록은 무시)
                    if (
                        day_records_sorted.iloc[-1]["기록유형"] != "경비시작"
                        and day_records_sorted.iloc[-1]["기록유형"] != "기타"
                    ):
                        last_record = day_records_sorted.iloc[-1]
                        last_record_time = (
                            f"{last_record['시간_시']:02d}:{last_record['시간_분']:02d}"
                        )
                        unclear_security_days.append(
                            {
                                "업무일": business_day,
                                "마지막기록시간": last_record_time,
                                "기록유형": last_record["기록유형"],
                                "문제": "마지막 기록이 '경비시작'이 아님",
                            }
                        )
                        print(
                            f"[의심데이터] {business_day} - 마지막 기록이 '경비시작'이 아닙니다 (유형: {last_record['기록유형']}, 시간: {last_record_time})"
                        )

                # 명확한 경비 기록이 아예 없는 경우도 의심 데이터로 분류
                if not has_release and not has_start and len(day_records) > 0:
                    unclear_security_days.append(
                        {
                            "업무일": business_day,
                            "기록수": len(day_records),
                            "문제": "명확한 경비해제/시작 기록 없음",
                        }
                    )
                    print(f"[의심데이터] {business_day} - 명확한 경비 기록 없음 (사용자 확인 필요)")

            # 의심 기록 저장
            self.unclear_security_days = unclear_security_days

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
            # 데이터가 3행부터 시작하고, 열 위치가 고정된 경우를 처리
            # 2열에 헤더가 있는 경우 확인
            if len(df) >= 2:
                # 필요하면 1-2행 제거 (헤더 및 SQL 실행 결과)
                df = df.iloc[2:].reset_index(drop=True)

            # 위치 기반으로 컬럼 매핑
            col_mapping = {
                "날짜": 6,  # G열 - 초과근무일자
                "시작시간": 7,  # H열 - 출근시간
                "종료시간": 8,  # I열 - 퇴근시간
                "이름": 3,  # D열 - 성명
            }

            # 인덱스 기반으로 컬럼 이름 만들기
            renamed_columns = {}
            for i in range(len(df.columns)):
                if i == 0:
                    renamed_columns[df.columns[i]] = "부서명"
                elif i == 1:
                    renamed_columns[df.columns[i]] = "직급"
                elif i == 2:
                    renamed_columns[df.columns[i]] = "개인식별번호"
                elif i == 3:
                    renamed_columns[df.columns[i]] = "성명"
                elif i == 4:
                    renamed_columns[df.columns[i]] = "현업여부"
                elif i == 5:
                    renamed_columns[df.columns[i]] = "휴일여부"
                elif i == 6:
                    renamed_columns[df.columns[i]] = "초과근무일자"
                elif i == 7:
                    renamed_columns[df.columns[i]] = "출근시간"
                elif i == 8:
                    renamed_columns[df.columns[i]] = "퇴근시간"
                elif i == 9:
                    renamed_columns[df.columns[i]] = "출근IP"
                elif i == 10:
                    renamed_columns[df.columns[i]] = "퇴근IP"
                elif i == 11:
                    renamed_columns[df.columns[i]] = "초과근무시간"
                elif i == 12:
                    renamed_columns[df.columns[i]] = "수당시간"
                elif i == 13:
                    renamed_columns[df.columns[i]] = "근무내용"
                else:
                    renamed_columns[df.columns[i]] = f"컬럼{i}"

            # 데이터프레임 컬럼 이름 변경
            df = df.rename(columns=renamed_columns)

            # 컬럼 이름 기반으로 매핑
            col_mapping = {
                "날짜": "초과근무일자",
                "시작시간": "출근시간",
                "종료시간": "퇴근시간",
                "이름": "성명",
            }

            # 날짜 필터링 적용
            start_date = self.start_date.date().toString("yyyy-MM-dd")
            end_date = self.end_date.date().toString("yyyy-MM-dd")

            # 날짜 데이터 정리 (YYYY-MM-DD 형식 고정)
            # 데이터프레임에서 초과근무일자 열이 문자열이면 datetime으로 변환
            df["날짜_datetime"] = pd.to_datetime(
                df[col_mapping["날짜"]], format="%Y-%m-%d", errors="coerce"
            )

            def parse_time(time_value, default_time):
                """시간 값을 파싱하여 time 객체로 반환합니다."""
                if pd.isna(time_value):
                    return default_time

                try:
                    if isinstance(time_value, str):
                        time_str = time_value.strip()
                        parts = time_str.split(":")
                        if len(parts) >= 2:
                            hour = int(parts[0])
                            minute = int(parts[1])
                            return time(hour, minute)
                        else:
                            print(f"시간 형식 오류 (HH:mm 형식이 아님): {time_str}")
                            return default_time
                    elif isinstance(time_value, datetime):
                        return time_value.time()
                    elif isinstance(time_value, time):
                        return time_value
                    else:
                        print(f"지원하지 않는 시간 형식: {type(time_value)}")
                        return default_time
                except Exception as e:
                    print(f"시간 파싱 오류: {str(e)}")
                    return default_time

            # 시간 데이터 유효성 검사
            # 시간 형식 검사 함수 (HH:mm 형식 고정)
            def is_valid_time_format(time_str):
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

            # 필터링 적용
            filtered_df = df
            if start_date and end_date:
                filtered_df = df[
                    (df["날짜_datetime"] >= start_date) & (df["날짜_datetime"] <= end_date)
                ]

            # 각 행의 데이터 유효성 확인을 위한 작업
            filtered_df["데이터_유효"] = (
                pd.notna(filtered_df[col_mapping["이름"]])
                & pd.notna(filtered_df[col_mapping["날짜"]])
                & pd.notna(filtered_df[col_mapping["시작시간"]])
                & pd.notna(filtered_df[col_mapping["종료시간"]])
                & pd.notna(filtered_df["날짜_datetime"])
            )

            # 유효한 데이터만 선택
            filtered_df = filtered_df[filtered_df["데이터_유효"]]

            # 초과근무 데이터 정리
            overtime_records = []

            # 정규 근무시간 설정 (9:00-18:00)
            regular_start = time(9, 0)
            regular_end = time(18, 0)

            for _, row in filtered_df.iterrows():
                try:
                    # 날짜 처리
                    work_date = row["날짜_datetime"].date()

                    # 출/퇴근 시간 누락 여부 플래그
                    has_missing_time = False
                    missing_time_fields = []

                    # 출근시간 처리 (HH:mm 고정 형식)
                    if pd.isna(row[col_mapping["시작시간"]]):
                        has_missing_time = True
                        missing_time_fields.append("출근시간")

                    start_time = parse_time(row[col_mapping["시작시간"]], regular_start)

                    if start_time is None:
                        # 시작 시간이 없거나 파싱 실패
                        has_missing_time = True
                        missing_time_fields.append("출근시간")
                        start_time = regular_start  # 일단 기본값 설정 (나중에 의심 데이터로 표시)

                    # 퇴근시간 처리 (HH:mm 고정 형식)
                    if pd.isna(row[col_mapping["종료시간"]]):
                        has_missing_time = True
                        missing_time_fields.append("퇴근시간")

                    end_time = parse_time(row[col_mapping["종료시간"]], regular_end)

                    if end_time is None:
                        # 종료 시간이 없거나 파싱 실패
                        has_missing_time = True
                        missing_time_fields.append("퇴근시간")
                        end_time = regular_end  # 일단 기본값 설정 (나중에 의심 데이터로 표시)

                    # 업무일 결정 (새벽 4시 기준)
                    business_date = work_date

                    # 자정 이후 새벽 4시 이전 근무는 전날 업무일로 계산
                    if end_time < time(4, 0):
                        business_date = work_date - timedelta(days=1)

                    # 휴일 여부 확인 (F열 데이터만 사용)
                    is_holiday = False

                    # F열(휴일여부) 데이터 확인
                    if "휴일여부" in df.columns and pd.notna(row["휴일여부"]):
                        holiday_value = str(row["휴일여부"]).strip().lower()
                        # "Y", "휴일", "공휴일", "토요일", "일요일" 등의 문자가 포함되어 있으면 휴일로 판단
                        holiday_keywords = ["y", "휴", "공휴", "토요일", "일요일"]
                        is_holiday = any(x in holiday_value for x in holiday_keywords)
                        matched_keywords = [x for x in holiday_keywords if x in holiday_value]
                        if matched_keywords:
                            print(
                                f"[휴일판단] 날짜: {business_date}, 휴일여부: {is_holiday}, F열 데이터: '{row.get('휴일여부', '데이터 없음')}', 일치 키워드: {matched_keywords}"
                            )
                        else:
                            print(
                                f"[평일판단] 날짜: {business_date}, 평일근무, F열 데이터: '{row.get('휴일여부', '데이터 없음')}'"
                            )
                    else:
                        # F열 데이터가 없거나 누락된 경우 기본값은 평일(False)로 설정
                        print(
                            f"[데이터없음] 날짜: {business_date}, 평일로 처리(기본값), F열 데이터: '{row.get('휴일여부', '데이터 없음')}'"
                        )

                    # 디버그 정보 출력
                    if is_holiday:
                        print(
                            f"[휴일처리] {business_date} - {row.get('성명', '이름없음')} - 휴일로 처리됨"
                        )

                    # 초과근무 여부 확인
                    # 휴일인 경우: 모든 시간이 초과근무 시간
                    # 평일인 경우: 시작 시간이 18시 이후이거나 종료 시간이 9시 이전인 경우만 초과근무
                    is_overtime = is_holiday or (
                        start_time >= regular_end or end_time <= regular_start
                    )

                    # 직원 이름 정보
                    employee_name = (
                        str(row[col_mapping["이름"]])
                        if pd.notna(row[col_mapping["이름"]])
                        else "Unknown"
                    )
                    # 부서명 정보 추가 (있는 경우)
                    department = (
                        str(row["부서명"]) if "부서명" in row and pd.notna(row["부서명"]) else ""
                    )  # 초과근무시간 정보 (있는 경우 사용)
                    overtime_hours = None
                    if "초과근무시간" in df.columns and pd.notna(row["초과근무시간"]):
                        try:
                            # 문자열이면 숫자로 변환 시도
                            if isinstance(row["초과근무시간"], str):
                                # 콤마, 공백 등 제거하고 숫자 변환
                                clean_str = row["초과근무시간"].replace(",", "").strip()
                                overtime_hours = float(clean_str)
                            else:
                                overtime_hours = float(row["초과근무시간"])
                        except:
                            print(f"초과근무시간 변환 실패: {row['초과근무시간']}")
                            pass

                    # 추가 정보 (근무내용)
                    work_description = ""
                    if "근무내용" in df.columns and pd.notna(row["근무내용"]):
                        work_description = str(row["근무내용"]).strip()

                    # 휴일인 경우 모든 시간을 초과근무로 처리
                    if is_holiday:
                        # 휴일 근무는 전체가 초과근무
                        overtime_start = start_time
                        overtime_end = end_time

                        overtime_records.append(
                            {
                                "업무일": business_date,
                                "날짜": work_date,
                                "시작시간": overtime_start,
                                "종료시간": overtime_end,
                                "초과근무유형": "휴일근무",
                                "직원명": employee_name,
                                "부서명": department,
                                "기록된_초과근무시간": overtime_hours,
                                "근무내용": work_description,
                                "휴일여부": True,
                            }
                        )
                    # 평일인 경우 정규 근무시간(9-18)을 제외한 시간만 초과근무로 처리
                    else:
                        # 업무 시간이 정규 근무시간(9-18)에 걸쳐있는 경우, 그 부분은 초과근무가 아님
                        if not is_overtime and (
                            start_time < regular_end and end_time > regular_start
                        ):
                            # 시작 시간이 9시 이전이면 초과근무 시작 부분 기록
                            if start_time < regular_start:
                                overtime_start = start_time
                                overtime_end = regular_start

                                overtime_records.append(
                                    {
                                        "업무일": business_date,
                                        "날짜": work_date,
                                        "시작시간": overtime_start,
                                        "종료시간": overtime_end,
                                        "초과근무유형": "조기출근",
                                        "직원명": employee_name,
                                        "부서명": department,
                                        "기록된_초과근무시간": overtime_hours,
                                        "근무내용": work_description,
                                        "휴일여부": False,
                                    }
                                )

                            # 종료 시간이 18시 이후면 초과근무 종료 부분 기록
                            if end_time > regular_end:
                                overtime_start = regular_end
                                overtime_end = end_time

                                overtime_records.append(
                                    {
                                        "업무일": business_date,
                                        "날짜": work_date,
                                        "시작시간": overtime_start,
                                        "종료시간": overtime_end,
                                        "초과근무유형": "야근",
                                        "직원명": employee_name,
                                        "부서명": department,
                                        "기록된_초과근무시간": overtime_hours,
                                        "근무내용": work_description,
                                        "휴일여부": False,
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

                            overtime_records.append(
                                {
                                    "업무일": business_date,
                                    "날짜": work_date,
                                    "시작시간": overtime_start,
                                    "종료시간": overtime_end,
                                    "초과근무유형": overtime_type,
                                    "직원명": employee_name,
                                    "부서명": department,
                                    "기록된_초과근무시간": overtime_hours,
                                    "근무내용": work_description,
                                    "휴일여부": False,
                                }
                            )

                except Exception as e:
                    print(f"초과근무 기록 처리 중 오류: {str(e)}")
                    error_desc = str(e)
                    # 오류 발생 데이터 저장
                    error_record = {
                        "업무일": (
                            business_date
                            if "business_date" in locals()
                            else work_date if "work_date" in locals() else "알 수 없음"
                        ),
                        "직원명": (
                            employee_name
                            if "employee_name" in locals()
                            else row.get(col_mapping["이름"], "알 수 없음")
                        ),
                        "오류내용": error_desc,
                        "원본데이터": {
                            k: str(v)
                            for k, v in row.items()
                            if k
                            in [
                                col_mapping["날짜"],
                                col_mapping["시작시간"],
                                col_mapping["종료시간"],
                                col_mapping["이름"],
                                "부서명",
                                "근무내용",
                            ]
                        },
                    }

                    # 오류 데이터 목록에 추가
                    if not hasattr(self, "error_records"):
                        self.error_records = []
                    self.error_records.append(error_record)
                    continue

            # 누락된 시간 정보가 있는 데이터 검사 및 의심 데이터로 추가
            missing_time_records = []
            for record in filtered_df.iterrows():
                row = record[1]
                if pd.isna(row[col_mapping["시작시간"]]) or pd.isna(row[col_mapping["종료시간"]]):
                    work_date = (
                        row["날짜_datetime"].date() if pd.notna(row["날짜_datetime"]) else None
                    )
                    employee_name = (
                        str(row[col_mapping["이름"]])
                        if pd.notna(row[col_mapping["이름"]])
                        else "Unknown"
                    )
                    department = (
                        str(row["부서명"]) if "부서명" in row and pd.notna(row["부서명"]) else ""
                    )

                    missing_fields = []
                    if pd.isna(row[col_mapping["시작시간"]]):
                        missing_fields.append("출근시간")
                    if pd.isna(row[col_mapping["종료시간"]]):
                        missing_fields.append("퇴근시간")

                    missing_time_records.append(
                        {
                            "업무일": work_date,
                            "직원명": employee_name,
                            "부서명": department,
                            "누락필드": ", ".join(missing_fields),
                            "원본데이터": {
                                k: str(v)
                                for k, v in row.items()
                                if k
                                in [
                                    col_mapping["날짜"],
                                    col_mapping["시작시간"],
                                    col_mapping["종료시간"],
                                    col_mapping["이름"],
                                    "부서명",
                                    "근무내용",
                                ]
                            },
                        }
                    )

            # 시간 누락 기록 저장
            self.missing_time_records = missing_time_records
            if missing_time_records:
                print(
                    f"[주의] {len(missing_time_records)}개의 출/퇴근 시간 누락 기록이 발견되었습니다."
                )

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
                print(
                    f"[의심데이터] {business_date} - {employee_name} - 경비 기록 없음 (사용자 확인 필요)"
                )
                # 경비 기록이 없어도 의심 데이터로 추가
                suspicious_reason = "해당 업무일에 경비 기록 없음"

                # 초과근무 기록에서 추가 정보 수집
                work_content = ""
                department = ""
                is_holiday = False

                for ovt_record in overtime_records:
                    if (
                        ovt_record["직원명"] == employee_name
                        and ovt_record["업무일"] == business_date
                    ):
                        if "부서명" in ovt_record and ovt_record["부서명"]:
                            department = ovt_record["부서명"]
                        if "근무내용" in ovt_record and ovt_record["근무내용"]:
                            work_content = ovt_record["근무내용"]
                        if "휴일여부" in ovt_record:
                            is_holiday = ovt_record["휴일여부"]
                        break

                suspicious_records.append(
                    {
                        "날짜": business_date,
                        "직원명": employee_name,
                        "부서명": department,
                        "초과근무시간": f"{overtime_start.strftime('%H:%M')}-{overtime_end.strftime('%H:%M')}",
                        "경비상태": "기록 없음",
                        "의심사유": suspicious_reason,
                        "근무내용": work_content,
                        "휴일여부": "휴일" if is_holiday else "평일",
                    }
                )
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

            # 경비 상태 변화 기록
            security_changes = []

            # 경비 상태 기록을 시간 순으로 정렬하고 상태 변화 추적
            for record in security_status:
                record_time = datetime.strptime(record["시간"], "%Y-%m-%d %H:%M")
                security_changes.append({"시간": record_time, "상태": record["상태"]})

            # 경비 변화가 없으면 다음 기록으로 넘어감
            if not security_changes:
                continue

            # 초기 상태 설정
            security_changes.sort(key=lambda x: x["시간"])

            # 의심 시간 구간 계산
            suspicious_intervals = []

            # 초과근무 시간과 경비 상태를 구간별로 비교하여 의심 시간대 계산
            # 경비 상태의 시간대별 구간 생성 (경비시작-경비해제 구간)
            security_periods = []

            # 경비 변화 상태에 따른 시간대 구간 생성
            if len(security_changes) > 0:
                # 먼저 첫 상태가 "해제"인 경우, 자정부터 첫 해제까지는 경비 활성화 상태로 간주
                if security_changes[0]["상태"] == "해제":
                    midnight = datetime.combine(overtime["날짜"], time(0, 0))
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
                            next_day = datetime.combine(
                                overtime["날짜"] + timedelta(days=1), time(0, 0)
                            )
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

            # 경비기록이 하나도 없는 경우 (상태 변화가 없는 경우)
            if len(security_periods) == 0:
                # 기본적으로 보수적인 접근: 경비 활성화 상태로 간주
                today_start = datetime.combine(overtime["날짜"], time(0, 0))
                tomorrow_start = datetime.combine(overtime["날짜"] + timedelta(days=1), time(0, 0))
                security_periods.append(
                    {
                        "시작": today_start,
                        "종료": tomorrow_start,
                        "상태": "시작",  # 경비 활성화 상태
                    }
                )

            # 초과근무 시간과 경비 활성화 시간대를 비교하여 의심 구간 계산
            suspicious_intervals = []

            # 각 경비 활성화 구간과 초과근무 시간 비교
            for period in security_periods:
                # 경비 활성화 상태인 경우만 검사
                if period["상태"] == "시작":
                    # 초과근무 시간이 경비 활성화 구간과 겹치는지 확인
                    if max(period["시작"], overtime_start_dt) < min(
                        period["종료"], overtime_end_dt
                    ):
                        # 겹치는 구간 계산
                        overlap_start = max(period["시작"], overtime_start_dt)
                        overlap_end = min(period["종료"], overtime_end_dt)
                        suspicious_intervals.append((overlap_start, overlap_end))

                        # 디버그 출력
                        print(
                            f"[의심기록] {business_date} - {employee_name} - 경비활성화({period['시작'].strftime('%H:%M:%S')}-{period['종료'].strftime('%H:%M:%S')}) 중 초과근무 발생({overlap_start.strftime('%H:%M:%S')}-{overlap_end.strftime('%H:%M:%S')})"
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

            # 의심 시간이 있으면 기록
            if total_suspicious_hours > 0:  # 1분이라도 의심 시간이 있으면 기록
                period_str = ", ".join(suspicious_periods)
                suspicious_reason = f"경비 작동 중 총 {total_suspicious_hours:.1f}시간 초과근무 기록 존재 ({period_str})"

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

                # 해당 직원의 초과근무 기록 중에서 부가 정보 찾기
                for ovt_record in overtime_records:
                    if (
                        ovt_record["직원명"] == employee_name
                        and ovt_record["업무일"] == business_date
                    ):
                        if "부서명" in ovt_record and ovt_record["부서명"]:
                            department = ovt_record["부서명"]
                        if "근무내용" in ovt_record and ovt_record["근무내용"]:
                            work_content = ovt_record["근무내용"]
                        break

                # 휴일 여부 파악
                is_holiday = False
                for ovt_record in overtime_records:
                    if (
                        ovt_record["직원명"] == employee_name
                        and ovt_record["업무일"] == business_date
                    ):
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
                    if (
                        change["상태"] == "시작"
                        and overtime_start_dt <= change["시간"] <= overtime_end_dt
                    ):
                        security_set_times.append(change["시간"].strftime("%H:%M:%S"))

                if security_set_times:
                    security_info += f" (경비설정시각: {', '.join(security_set_times)})"

                suspicious_records.append(
                    {
                        "날짜": business_date,
                        "직원명": employee_name,
                        "부서명": department,
                        "초과근무시간": f"{overtime_start.strftime('%H:%M')}-{overtime_end.strftime('%H:%M')}",
                        "경비상태": security_info,
                        "의심사유": suspicious_reason,
                        "근무내용": work_content,
                        "휴일여부": "휴일" if is_holiday else "평일",
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

        # 테이블 컬럼 수 조정 (추가 정보를 위해)
        self.table.setColumnCount(
            8
        )  # 날짜, 직원명, 부서명, 초과근무시간, 경비상태, 의심사유, 근무내용, 휴일여부
        self.table.setHorizontalHeaderLabels(
            [
                "날짜",
                "직원명",
                "부서명",
                "초과근무시간",
                "경비상태",
                "의심사유",
                "근무내용",
                "휴일여부",
            ]
        )

        # 테이블에 행 추가
        for i, record in enumerate(suspicious_records):
            self.table.insertRow(i)

            # 날짜
            date_item = QTableWidgetItem(
                record["날짜"].strftime("%Y-%m-%d")
                if hasattr(record["날짜"], "strftime")
                else str(record["날짜"])
            )
            self.table.setItem(i, 0, date_item)

            # 직원명
            name_item = QTableWidgetItem(str(record["직원명"]))
            self.table.setItem(i, 1, name_item)

            # 부서명
            dept_item = QTableWidgetItem(str(record.get("부서명", "")))
            self.table.setItem(i, 2, dept_item)

            # 초과근무 시간
            time_item = QTableWidgetItem(str(record["초과근무시간"]))
            self.table.setItem(i, 3, time_item)

            # 경비 상태
            security_item = QTableWidgetItem(str(record["경비상태"]))
            self.table.setItem(i, 4, security_item)

            # 의심 사유
            reason_item = QTableWidgetItem(str(record["의심사유"]))
            self.table.setItem(i, 5, reason_item)

            # 근무내용
            content_item = QTableWidgetItem(str(record.get("근무내용", "")))
            self.table.setItem(i, 6, content_item)

            # 휴일여부
            holiday_item = QTableWidgetItem(str(record.get("휴일여부", "평일")))
            self.table.setItem(i, 7, holiday_item)

    def export_results(self):
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
        """의심 기록만 엑셀로 내보냅니다."""
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OvertimeAnalyzer()
    window.show()
    sys.exit(app.exec_())
