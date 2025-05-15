# 로컬 개발 환경 설정 가이드

이 가이드는 로컬에서 초과근무 분석기를 개발하고 실행하기 위한 설정 방법을 안내합니다.

## 가상환경 설정 (macOS/Linux)

1. 프로젝트 루트 디렉토리에서 가상환경 생성:

   ```
   python3 -m venv venv
   ```

2. 가상환경 활성화:

   ```
   source venv/bin/activate
   ```

3. 필요한 패키지 설치:
   ```
   pip install -r requirements.txt
   ```

## 애플리케이션 실행

### macOS에서 실행

macOS에서는 PyQt5 플러그인 경로 설정이 필요할 수 있습니다:

1. 아래 명령어로 실행:

   ```
   ./run.sh
   ```

2. 또는 직접 환경변수 설정 후 실행:
   ```
   export QT_QPA_PLATFORM_PLUGIN_PATH="$(pwd)/venv/lib/python3.13/site-packages/PyQt5/Qt5/plugins"
   python app.py
   ```

### Windows에서 실행

1. 명령 프롬프트에서 가상환경 활성화:

   ```
   venv\Scripts\activate
   ```

2. 앱 실행:
   ```
   python app.py
   ```
   또는 run.bat 파일 실행

## 디버깅 팁

1. 가상환경이 활성화되어 있는지 확인
2. PyQt5가 제대로 설치되었는지 확인
3. macOS에서 문제가 발생하면 QT_QPA_PLATFORM_PLUGIN_PATH 환경변수 설정
