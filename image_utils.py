import os
from pathlib import Path
from PIL import Image

# 지원하는 이미지 파일 확장자
SUPPORTED_IMAGE_FORMATS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif', '.webp'}

def is_image_file(file_path):
    """파일 확장자를 기준으로 이미지 파일 여부를 확인합니다."""
    if not file_path or not isinstance(file_path, str):
        return False

    try:
        # 확장자 체크만으로 우선 판단 (성능 및 유연성)
        ext = Path(file_path).suffix.lower()
        return ext in SUPPORTED_IMAGE_FORMATS
    except:
        return False

def validate_image_path(file_path):
    """이미지 파일의 경로 유효성 및 실제 이미지 여부를 검증합니다."""
    if not file_path or not isinstance(file_path, str):
        return False, "파일 경로가 비어있습니다."

    if not os.path.exists(file_path):
        return False, f"파일이 존재하지 않습니다: {file_path}"

    if not os.path.isfile(file_path):
        return False, f"파일이 아닙니다: {file_path}"

    # 실제 이미지로 열 수 있는지 확인 (verify 대신 open 시도)
    try:
        with Image.open(file_path) as img:
            img.load() # 실제 데이터 로드 시도
        return True, "유효한 이미지 파일입니다."
    except Exception as e:
        return False, f"이미지 파일을 읽을 수 없습니다: {e}"

    # 파일 크기 체크 (10MB 제한)
    file_size = os.path.getsize(file_path)
    if file_size > 10 * 1024 * 1024:
        return False, f"파일 크기가 너무 큽니다 (10MB 제한): {file_size / (1024*1024):.1f}MB"

    return True, "유효한 이미지 파일입니다."

def get_image_display_name(file_path):
    """이미지 파일의 표시 이름을 반환합니다."""
    if not file_path:
        return ""
    return f"📷 {Path(file_path).name}"

def normalize_image_path(file_path):
    """이미지 경로를 정규화합니다."""
    if not file_path:
        return ""
    return os.path.abspath(file_path)