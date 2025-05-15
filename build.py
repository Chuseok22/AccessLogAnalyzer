import PyInstaller.__main__
import os
import sys
import shutil
import argparse

# 인코딩 설정
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

# 경로 설정
script_path = os.path.abspath(os.path.dirname(__file__))
main_script = os.path.join(script_path, "app.py")
dist_path = os.path.join(script_path, "dist")

# 명령행 인수 파싱
parser = argparse.ArgumentParser(description="초과근무 분석기 빌드 스크립트")
parser.add_argument("--clean-build", action="store_true", help="빌드 전 이전 빌드 파일 삭제")
parser.add_argument("--console", action="store_true", help="콘솔 창 표시")
args = parser.parse_args()

# 빌드 폴더 정리
if args.clean_build and os.path.exists(dist_path):
    print(f"기존 dist 폴더 삭제: {dist_path}")
    shutil.rmtree(dist_path)
    os.makedirs(dist_path)

# PyInstaller 옵션
pyinstaller_options = [
    main_script,
    "--name=OvertimeAnalyzer",
    "--onefile",
    "--clean",
    "--distpath={}".format(dist_path),
    "--add-data={}".format(os.path.join(script_path, "requirements.txt") + os.path.pathsep + "."),
    "--paths={}".format(script_path),
    "--paths={}".format(os.path.join(script_path, "src")),
    "--hidden-import=src",
    "--hidden-import=src.ui.analyzer_ui",
    "--hidden-import=src.services.analyzer_service",
    "--hidden-import=src.services.security_processor",
    "--hidden-import=src.services.overtime_processor",
    "--hidden-import=src.models.data_models",
    "--hidden-import=src.utils.date_utils",
    "--hidden-import=src.utils.file_utils",
    "--noupx",
]

# 콘솔 창 설정
if not args.console:
    pyinstaller_options.append("--windowed")
else:
    pyinstaller_options.append("--console")

print(f"빌드 명령어: {main_script}")
PyInstaller.__main__.run(pyinstaller_options)

# 빌드 결과 확인
exe_file = os.path.join(dist_path, "OvertimeAnalyzer.exe")
if os.path.exists(exe_file):
    print(f"빌드 성공: {exe_file}")
else:
    print(f"빌드 실패: {exe_file} 파일을 찾을 수 없습니다.")
    sys.exit(1)

print("빌드가 완료되었습니다!")
