#!/usr/bin/env python
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

from src.overtime_analyzer.main import main

if __name__ == "__main__":
    main()
