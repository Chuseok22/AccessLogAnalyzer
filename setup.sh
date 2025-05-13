#!/bin/bash

# 스크립트 위치 확인
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR"

echo "출입 기록 분석기 초기 설정을 시작합니다..."

# Python이 설치되어 있는지 확인
if ! command -v python3 &> /dev/null; then
    echo "Python 3가 설치되어 있지 않습니다. Python을 설치해주세요."
    exit 1
fi

# 가상 환경 생성
echo "가상 환경을 생성합니다..."
python3 -m venv venv

# 가상 환경 활성화
echo "가상 환경을 활성화합니다..."
source venv/bin/activate

# 필요한 패키지 설치
echo "필요한 패키지를 설치합니다..."
pip install --upgrade pip
pip install -r requirements.txt

echo "설정이 완료되었습니다!"
echo "프로그램을 실행하려면 다음 명령어를 입력하세요:"
echo "source venv/bin/activate && python access_log_analyzer.py"
