import PyInstaller.__main__
import os
import sys

# 현재 스크립트 경로
script_path = os.path.abspath(os.path.dirname(__file__))
main_script = os.path.join(script_path, "access_log_analyzer.py")

# PyInstaller 옵션
PyInstaller.__main__.run([
    main_script,
    "--name=출입기록분석기",
    "--onefile",
    "--windowed",
    "--icon=NONE",
    "--clean",
    "--add-data={}".format(os.path.join(script_path, "requirements.txt") + 
                           os.path.pathsep + "."),
    "--noupx",
])

print("빌드가 완료되었습니다!")
