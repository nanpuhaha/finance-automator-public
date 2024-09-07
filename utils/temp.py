import shutil
from pathlib import Path

TEMP_DIR = Path(__file__).parent.parent / "temp"
TEMP_DIR.mkdir(exist_ok=True, parents=True)


# TEMP_DIR 비우기 함수
def clear_temp_dir(temp_dir: Path = TEMP_DIR):
    """TEMP 디렉토리의 모든 파일 삭제"""
    if temp_dir.exists() and temp_dir.is_dir():
        print(f"TEMP 디렉토리 비우기: {temp_dir}")
        shutil.rmtree(temp_dir)
        print(f"shutil.rmtree 이후 {temp_dir.exists() = }")
        temp_dir.mkdir(exist_ok=True)
