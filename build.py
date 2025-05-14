import PyInstaller.__main__
import os
import sys
import shutil
import argparse

# Window CP1252 환경 stdout 인코딩 설정
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

# 현재 스크립트 경로
script_path = os.path.abspath(os.path.dirname(__file__))
main_script = os.path.join(script_path, "run_analyzer.py")

# dist 폴더 확인 및 생성
dist_path = os.path.join(script_path, "dist")
if not os.path.exists(dist_path):
    os.makedirs(dist_path)

# 명령행 인수 파싱
parser = argparse.ArgumentParser(description="초과근무 분석기 빌드 스크립트")
parser.add_argument("--console", action="store_true", help="콘솔 창과 함께 실행")
parser.add_argument("--clean-build", action="store_true", help="빌드 전 이전 빌드 파일 삭제")
args = parser.parse_args()

# 빌드 전 정리
if args.clean_build:
    build_path = os.path.join(script_path, "build")
    if os.path.exists(build_path):
        print(f"기존 build 폴더 삭제: {build_path}")
        shutil.rmtree(build_path)

    if os.path.exists(dist_path):
        print(f"기존 dist 폴더 삭제: {dist_path}")
        shutil.rmtree(dist_path)
        os.makedirs(dist_path)

# PyInstaller 옵션
pyinstaller_args = [
    main_script,
    "--name=초과근무분석기",
    "--onefile",
    "--clean",
    "--distpath={}".format(dist_path),  # 명시적으로 dist 경로 지정
    "--add-data={}".format(os.path.join(script_path, "requirements.txt") + os.path.pathsep + "."),
    "--noupx",
]

# 콘솔 창 표시 여부
if not args.console:
    pyinstaller_args.append("--windowed")

# PyInstaller 실행
print(f"빌드 명령어: {' '.join(pyinstaller_args)}")
PyInstaller.__main__.run(pyinstaller_args)

# 빌드된 파일 경로 확인
try:
    # 운영체제에 맞게 확장자 결정
    if sys.platform.startswith("win"):
        exe_extension = ".exe"
    elif sys.platform.startswith("darwin"):
        exe_extension = ""  # macOS는 확장자가 없음
    else:  # Linux
        exe_extension = ""

    exe_file = os.path.join(dist_path, f"출입기록분석기{exe_extension}")

    if os.path.exists(exe_file):
        print(f"빌드 성공: {exe_file}")
    else:
        print(f"빌드 실패: {exe_file} 파일을 찾을 수 없습니다.")
        sys.exit(1)
except Exception as e:
    print(f"파일 확인 중 오류 발생: {str(e)}")
    sys.exit(1)

print("빌드가 완료되었습니다!")
