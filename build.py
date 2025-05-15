import PyInstaller.__main__
import os
import sys
import shutil
import argparse
import PyQt5

# PyQt5 플러그인 경로 확인
pyqt_path = os.path.dirname(PyQt5.__file__)
print(f"PyQt5 경로: {pyqt_path}")
qt_plugins_path = os.path.join(pyqt_path, "Qt5", "plugins")
print(f"Qt 플러그인 경로: {qt_plugins_path}")

# 만약 기본 경로에 플러그인이 없으면 대체 경로 확인
if not os.path.exists(qt_plugins_path):
    # 대체 경로 1 - Qt5 대신 Qt 디렉토리
    alt_plugins_path1 = os.path.join(pyqt_path, "Qt", "plugins")
    if os.path.exists(alt_plugins_path1):
        qt_plugins_path = alt_plugins_path1
        print(f"대체 플러그인 경로 사용: {qt_plugins_path}")
    else:
        # macOS에서 사용되는 다른 일반적인 경로
        for qt_ver in ["", "5"]:
            possible_path = os.path.join(sys.prefix, "plugins" + qt_ver)
            if os.path.exists(possible_path):
                qt_plugins_path = possible_path
                print(f"시스템 플러그인 경로 사용: {qt_plugins_path}")
                break
        else:
            print("경고: PyQt5 플러그인 경로를 찾을 수 없습니다. 빌드가 실패할 수 있습니다.")

# Window CP1252 환경 stdout 인코딩 설정
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

# 현재 스크립트 경로
script_path = os.path.abspath(os.path.dirname(__file__))
main_script = os.path.join(script_path, "overtime_analyzer.py")

# 메인 스크립트 존재 확인
if not os.path.exists(main_script):
    print(f"메인 스크립트가 존재하지 않습니다: {main_script}")
    sys.exit(1)

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
# 운영체제에 따른 설정
if sys.platform.startswith("win"):
    # Windows용 설정
    output_mode = "--onefile"  # Windows에서는 단일 파일로 빌드
else:
    # macOS, Linux용 설정
    output_mode = "--onedir"  # macOS에서는 onedir 모드가 더 안정적

pyinstaller_args = [
    main_script,
    "--name=OvertimeAnalyzer",
    output_mode,
    "--clean",
    "--distpath={}".format(dist_path),  # 명시적으로 dist 경로 지정
    "--add-data={}".format(os.path.join(script_path, "requirements.txt") + os.path.pathsep + "."),
    "--paths={}".format(script_path),  # 프로젝트 루트 디렉토리 추가
    "--paths={}".format(os.path.join(script_path, "src")),  # src 디렉토리도 추가
    "--hidden-import=src",  # src 패키지 자체도 추가
    "--hidden-import=src.overtime_analyzer",  # 모듈 import 명시적 추가
    "--hidden-import=src.overtime_analyzer.ui.analyzer_ui",  # UI 모듈 명시적 추가
    "--hidden-import=src.overtime_analyzer.services.analyzer_service",
    "--hidden-import=src.overtime_analyzer.services.security_processor",
    "--hidden-import=src.overtime_analyzer.services.overtime_processor",
    "--hidden-import=src.overtime_analyzer.models.data_models",
    "--hidden-import=src.overtime_analyzer.utils.date_utils",
    "--hidden-import=src.overtime_analyzer.utils.file_utils",
    "--hidden-import=PyQt5",
    "--hidden-import=PyQt5.QtCore",
    "--hidden-import=PyQt5.QtGui",
    "--hidden-import=PyQt5.QtWidgets",
    "--noupx",
]

# Qt 플러그인 경로가 확인되었다면 추가
if "qt_plugins_path" in locals() and os.path.exists(qt_plugins_path):
    pyinstaller_args.append(
        "--add-binary={}{}{}".format(qt_plugins_path, os.path.sep, "PyQt5/Qt/plugins")
    )

# 콘솔 창 표시 여부
if not args.console:
    pyinstaller_args.append("--windowed")

# PyInstaller 실행
print(f"빌드 명령어: {' '.join(pyinstaller_args)}")
try:
    PyInstaller.__main__.run(pyinstaller_args)
except Exception as e:
    print(f"PyInstaller 실행 중 오류 발생: {str(e)}")
    sys.exit(1)

# 빌드된 파일 경로 확인
try:
    # 운영체제에 맞게 확장자 결정
    if sys.platform.startswith("win"):
        exe_extension = ".exe"
        exe_name = "OvertimeAnalyzer"
        exe_file = os.path.join(dist_path, f"{exe_name}{exe_extension}")
    elif sys.platform.startswith("darwin"):
        # macOS에서는 onedir 모드에서 .app 패키지로 생성됨
        exe_name = "OvertimeAnalyzer"
        exe_file = os.path.join(dist_path, f"{exe_name}.app")
    else:  # Linux
        exe_extension = ""
        exe_name = "OvertimeAnalyzer"
        exe_file = os.path.join(dist_path, f"{exe_name}{exe_extension}")

    if os.path.exists(exe_file):
        print(f"빌드 성공: {exe_file}")
    else:
        # 대체 경로 확인 (Windows onefile 모드)
        alt_file = os.path.join(dist_path, "OvertimeAnalyzer", "OvertimeAnalyzer.exe")
        if os.path.exists(alt_file):
            print(f"빌드 성공: {alt_file}")
        else:
            print(f"빌드 실패: {exe_file} 파일을 찾을 수 없습니다.")
            sys.exit(1)
except Exception as e:
    print(f"파일 확인 중 오류 발생: {str(e)}")
    sys.exit(1)

print("빌드가 완료되었습니다!")
