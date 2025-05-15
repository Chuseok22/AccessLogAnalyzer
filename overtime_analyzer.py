#!/usr/bin/env python3
"""
초과근무 분석기를 실행하기 위한 스크립트입니다.
이 스크립트는 프로젝트의 루트 디렉토리에서 실행해야 합니다.
"""
import os
import sys

# 경로를 추가하여 모듈 가져오기 문제 해결
script_dir = os.path.abspath(os.path.dirname(__file__))
src_dir = os.path.join(script_dir, 'src')
if os.path.exists(src_dir):
    sys.path.insert(0, script_dir)
    sys.path.insert(0, src_dir)

try:
    # 여러 방법으로 메인 모듈 가져오기 시도
    try:
        from src.overtime_analyzer.main import main
    except ImportError:
        try:
            from overtime_analyzer.main import main
        except ImportError:
            # 마지막 대안 - 직접 main 함수 구현
            def main():
                from PyQt5.QtWidgets import QApplication
                # 여러 방법으로 UI 모듈 가져오기 시도
                try:
                    from src.overtime_analyzer.ui.analyzer_ui import AnalyzerUI
                except ImportError:
                    from overtime_analyzer.ui.analyzer_ui import AnalyzerUI
                
                app = QApplication(sys.argv)
                window = AnalyzerUI()
                window.show()
                return app.exec()
except ImportError as e:
    print(f"패키지 로드 실패: {e}", file=sys.stderr)
    sys.exit(1)

if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print(f"애플리케이션 실행 중 오류 발생: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)