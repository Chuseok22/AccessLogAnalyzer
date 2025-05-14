import PyInstaller.__main__
import os
import sys
import shutil

# Window CP1252 환경 stdout 인코딩 설정
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

# 현재 스크립트 경로
script_path = os.path.abspath(os.path.dirname(__file__))
main_script = os.path.join(script_path, "overtime_analyzer.py")

# dist 폴더 확인 및 생성
dist_path = os.path.join(script_path, "dist")
if not os.path.exists(dist_path):
    os.makedirs(dist_path)

# PyInstaller 옵션
PyInstaller.__main__.run(
    [
        main_script,
        "--name=overtime_analyzer",  # 임시로 영문 이름 사용
        "--onefile",
        "--windowed",
        "--icon=NONE",
        "--clean",
        "--distpath={}".format(dist_path),  # 명시적으로 dist 경로 지정
        "--add-data={}".format(
            os.path.join(script_path, "requirements.txt") + os.path.pathsep + "."
        ),
        "--noupx",
    ]
)

# 빌드된 파일 경로 확인
try:
    exe_file = os.path.join(dist_path, "overtime_analyzer.exe")
    if os.path.exists(exe_file):
        print(f"빌드 성공: {exe_file}")
    else:
        print(f"빌드 실패: {exe_file} 파일을 찾을 수 없습니다.")
        sys.exit(1)
except Exception as e:
    print(f"파일 확인 중 오류 발생: {str(e)}")
    sys.exit(1)

print("빌드가 완료되었습니다!")
