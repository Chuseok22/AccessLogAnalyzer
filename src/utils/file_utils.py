"""
파일 관련 유틸리티 함수를 제공합니다.
"""

import os
import pandas as pd
from typing import Tuple, Optional


def load_excel_file(file_path: str) -> Optional[pd.DataFrame]:
    """
    엑셀 파일을 로드하여 데이터프레임으로 반환합니다.

    Args:
        file_path: 엑셀 파일 경로

    Returns:
        pandas DataFrame 또는 None (파일 로드 실패 시)

    Raises:
        ValueError: 지원하지 않는 파일 형식인 경우
        FileNotFoundError: 파일을 찾을 수 없는 경우
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    # 확장자에 따라 적절한 엔진 사용
    file_ext = os.path.splitext(file_path)[1].lower()

    try:
        if file_ext == ".xls":
            return pd.read_excel(file_path, engine="xlrd")
        elif file_ext == ".xlsx":
            return pd.read_excel(file_path, engine="openpyxl")
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {file_ext}")
    except Exception as e:
        print(f"엑셀 파일 로드 중 오류 발생: {str(e)}")
        return None


def save_excel_file(df: pd.DataFrame, file_path: str) -> bool:
    """
    데이터프레임을 엑셀 파일로 저장합니다.

    Args:
        df: 저장할 데이터프레임
        file_path: 저장할 파일 경로

    Returns:
        bool: 저장 성공 여부
    """
    try:
        # 경로 디렉토리가 없으면 생성
        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        # 확장자에 따라 적절한 엔진 사용
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext == ".xls":
            df.to_excel(file_path, engine="xlwt", index=False)
        elif file_ext == ".xlsx":
            df.to_excel(file_path, engine="openpyxl", index=False)
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {file_ext}")

        return True
    except Exception as e:
        print(f"엑셀 파일 저장 중 오류 발생: {str(e)}")
        return False
