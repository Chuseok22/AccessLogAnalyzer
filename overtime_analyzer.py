#!/usr/bin/env python3
"""
초과근무 분석기를 실행하기 위한 스크립트입니다.
이 스크립트는 프로젝트의 루트 디렉토리에서 실행해야 합니다.
"""
import os
import sys

# 현재 스크립트의 절대 경로
script_dir = os.path.dirname(os.path.abspath(__file__))

# 프로젝트 루트 디렉토리를 Python 경로에 추가
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# src 디렉토리 경로를 명시적으로 추가 (PyInstaller 호환성 개선)
src_dir = os.path.join(script_dir, "src")
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

try:
    # 먼저 상대 경로로 import 시도
    from src.overtime_analyzer.main import main
except ImportError:
    try:
        # 실패하면 직접 모듈로 import 시도
        import overtime_analyzer.main

        main = overtime_analyzer.main.main
    except ImportError:
        # 마지막 시도: PyInstaller 환경에서는 sys.frozen 속성이 존재함
        if getattr(sys, "frozen", False):
            # 직접 main 함수 구현 (PyInstaller 환경에서 마지막 대안)
            def main():
                print("PyInstaller 환경에서 직접 main 함수 실행")
                from PyQt5.QtWidgets import QApplication

                # src 경로에서 애플리케이션 UI 클래스 가져오기 시도
                try:
                    from src.overtime_analyzer.ui.analyzer_ui import AnalyzerUI
                except ImportError:
                    try:
                        from overtime_analyzer.ui.analyzer_ui import AnalyzerUI
                    except ImportError:
                        # 모든 방법이 실패한 경우
                        print("오류: AnalyzerUI 모듈을 로드할 수 없습니다.")
                        sys.exit(1)

                app = QApplication(sys.argv)
                window = AnalyzerUI()
                window.show()
                sys.exit(app.exec_())


if __name__ == "__main__":
    main()
