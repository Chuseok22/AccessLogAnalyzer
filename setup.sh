#!/bin/bash

# 스크립트 위치 확인
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR" || exit 1

echo "사무실 출입 및 초과근무 분석 도구 초기 설정을 시작합니다..."

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
echo "프로그램을 실행하시겠습니까?"
echo "1. 출입 기록 분석기"
echo "2. 초과근무 분석기"
echo "3. 나중에 실행"
read -p "선택 (1/2/3): " run_choice

if [ "$run_choice" = "1" ]; then
    python access_log_analyzer.py
elif [ "$run_choice" = "2" ]; then
    python overtime_analyzer.py
else
    echo ""
    echo "나중에 실행하려면:"
    echo "- 출입 기록 분석: './run.sh' 실행"
    echo "- 초과근무 분석: './run_overtime.sh' 실행"
fi
