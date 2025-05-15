#!/bin/bash

# 스크립트 위치 확인
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR" || exit 1

# 가상 환경 활성화
source venv/bin/activate

# 프로그램 실행
python app.py