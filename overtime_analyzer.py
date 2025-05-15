#!/usr/bin/env python3
"""
초과근무 분석기를 실행하기 위한 스크립트입니다.
이 스크립트는 프로젝트의 루트 디렉토리에서 실행해야 합니다.
"""
import os
import sys
import traceback

# 현재 스크립트 경로 및 src 경로를 설정
script_dir = os.path.abspath(os.path.dirname(__file__))
sys.path.insert(0, script_dir)  # 현재 디렉토리를 경로에 추가
src_dir = os.path.join(script_dir, 'src')
if os.path.exists(src_dir):
    sys.path.insert(0, src_dir)  # src 디렉토리를 경로에 추가

# 이 부분이 중요: 모든 필요한 기능을 직접 이 파일에서 import
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QFileDialog
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QGroupBox, QWidget, QTabWidget
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox
from PyQt5.QtWidgets import QDateEdit, QProgressBar, QSplitter
from PyQt5.QtCore import Qt, QDate
import pandas as pd
import numpy as np
from datetime import datetime, date, time, timedelta
import openpyxl
import os
import sys
import traceback

# 프로그램의 모든 모듈 import
try:
    # services
    from src.overtime_analyzer.services.analyzer_service import AnalyzerService
    from src.overtime_analyzer.services.security_processor import SecurityProcessor
    from src.overtime_analyzer.services.overtime_processor import OvertimeProcessor
    
    # models
    from src.overtime_analyzer.models.data_models import SecurityRecord, SuspiciousRecord
    
    # utils
    from src.overtime_analyzer.utils.date_utils import calculate_business_date
    from src.overtime_analyzer.utils.file_utils import save_to_excel
    
    # ui
    from src.overtime_analyzer.ui.analyzer_ui import AnalyzerUI
    
except ImportError:
    try:
        # 절대 경로 import
        from overtime_analyzer.services.analyzer_service import AnalyzerService
        from overtime_analyzer.services.security_processor import SecurityProcessor
        from overtime_analyzer.services.overtime_processor import OvertimeProcessor
        from overtime_analyzer.models.data_models import SecurityRecord, SuspiciousRecord
        from overtime_analyzer.utils.date_utils import calculate_business_date
        from overtime_analyzer.utils.file_utils import save_to_excel
        from overtime_analyzer.ui.analyzer_ui import AnalyzerUI
        
    except ImportError as e:
        print(f"모듈 가져오기 오류: {e}")
        traceback.print_exc()
        sys.exit(1)

def main():
    """애플리케이션 메인 함수"""
    app = QApplication(sys.argv)
    window = AnalyzerUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"애플리케이션 실행 중 오류 발생: {e}")
        traceback.print_exc()
        sys.exit(1)