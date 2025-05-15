#!/usr/bin/env python3
"""
초과근무 분석기를 실행하기 위한 스크립트입니다.
"""
import os
import sys

# 현재 디렉토리와 src 디렉토리를 경로에 추가
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)
sys.path.insert(0, os.path.join(script_dir, "src"))

from PyQt5.QtWidgets import QApplication
from src.ui.analyzer_ui import AnalyzerUI


def main():
    app = QApplication(sys.argv)
    window = AnalyzerUI()
    window.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
