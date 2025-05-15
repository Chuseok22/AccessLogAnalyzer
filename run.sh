#!/bin/bash

# 스크립트 위치 확인
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR" || exit 1

# 가상 환경 활성화
source venv/bin/activate

# PyQt5 플러그인 경로 설정
export QT_QPA_PLATFORM_PLUGIN_PATH="$SCRIPT_DIR/venv/lib/python3.13/site-packages/PyQt5/Qt5/plugins"

# 프로그램 실행
python app.py
