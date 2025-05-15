#!/usr/bin/env python
"""
초과근무 분석기 - 메인 애플리케이션 진입점

이 스크립트는 초과근무 분석기의 주요 진입점입니다.
GUI 애플리케이션을 시작하고 초기화합니다.
"""

import sys
from PyQt5.QtWidgets import QApplication
from overtime_analyzer.ui.analyzer_ui import AnalyzerUI


def main():
    """
    애플리케이션의 메인 함수
    """
    app = QApplication(sys.argv)
    window = AnalyzerUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
