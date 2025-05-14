#!/bin/bash

# 스크립트 위치 확인
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR" || exit 1

# 가상 환경 활성화 (존재하는 경우)
if [ -d "venv" ] && [ -f "venv/bin/activate" ]; then
  source venv/bin/activate
  echo "가상 환경이 활성화되었습니다."
else
  echo "가상 환경이 없습니다. 시스템 Python을 사용합니다."
fi

# 프로그램 실행
python overtime_analyzer.py
