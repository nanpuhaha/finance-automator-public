import io
import os
import shutil
from pathlib import Path
from zipfile import BadZipFile

import msoffcrypto
import openpyxl
import pandas as pd

PASSWORD = "1234"


def is_encrypted(path: Path) -> bool:
    """
    Check if the Excel file is encrypted.

    Args:
        path (Path): Path to the Excel file.

    Returns:
        bool: True if the Excel file is encrypted, False otherwise.
    """

    try:
        _ = pd.read_excel(path, engine='openpyxl')
    except BadZipFile as err:
        # zipfile.BadZipFile: File is not a zip file
        return True
    return False


def is_locked(path: Path) -> bool:
    """
    Check if the Excel file is locked.

    Args:
        path (Path): Path to the Excel file.

    Returns:
        bool: True if the Excel file is locked, False otherwise.
    """
    try:
        _ = openpyxl.load_workbook(filename=path)
    except BadZipFile as err:
        # zipfile.BadZipFile: File is not a zip file
        return True
    return False


def remove_password(path: Path, output_path: Path, password: str = PASSWORD) -> None:
    """
    Remove password from Excel file.

    Args:
        path (Path): Path to the Excel file.
        password (str): Password to decrypt the Excel file.
    """
    # 파일 사이즈 작음
    decrypted_workbook = io.BytesIO()

    with open(path, 'rb') as encrypted:
        file = msoffcrypto.OfficeFile(encrypted)
        file.load_key(password=password)
        file.decrypt(decrypted_workbook)

    workbook = openpyxl.load_workbook(filename=decrypted_workbook)
    workbook.save(output_path)


def unlock_password(path: Path, output_path: Path, password: str = PASSWORD) -> None:
    # 파일 사이즈 쬐끔 큼
    with open(path, "rb") as encrypted:
        file = msoffcrypto.OfficeFile(encrypted)
        file.load_key(password=password)

        with open(output_path, "wb") as f:
            file.decrypt(f)


def unlock_if_locked(path: Path, password: str = PASSWORD) -> None:
    if is_locked(path):
        output_path = path.with_stem(path.stem + "_unlocked")
        unlock_password(path, output_path, password)
        os.remove(path)
        shutil.move(output_path, path)
