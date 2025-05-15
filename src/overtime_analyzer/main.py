#!/usr/bin/env python
"""
초과근무 분석기 - 메인 애플리케이션 진입점

이 스크립트는 초과근무 분석기의 주요 진입점입니다.
GUI 애플리케이션을 시작하고 초기화합니다.
"""

import sys
from PyQt5.QtWidgets import QApplication

# 여러 import 방식 시도 - PyInstaller 호환성을 위한 처리
try:
    # 먼저 상대 경로 import 시도
    from .ui.analyzer_ui import AnalyzerUI
except ImportError:
    try:
        # 실패하면 절대 경로 import 시도
        from src.overtime_analyzer.ui.analyzer_ui import AnalyzerUI
    except ImportError:
        # 마지막으로 단순 모듈명 import 시도 (PyInstaller 환경용)
        from overtime_analyzer.ui.analyzer_ui import AnalyzerUI


def main():
    """
    애플리케이션의 메인 함수
    """
    app = QApplication(sys.argv)
    window = AnalyzerUI()
    window.show()
    sys.exit(app.exec())  # sys.exit 추가하여 정상 종료 보장


if __name__ == "__main__":
    main()