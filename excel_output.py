"""Automatic Excel export and optional local open/print actions."""

import os
from io import BytesIO
from pathlib import Path
from zipfile import is_zipfile


def save_generated_workbook(xlsx_bytes, project_dir):
    # Validates the generated Excel content before writing the workbook.
    if not xlsx_bytes or not is_zipfile(BytesIO(xlsx_bytes)):
        raise ValueError("The generated Excel workbook is empty or invalid.")

    # Creates a dedicated output folder within the project.
    output_dir = Path(project_dir).resolve() / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    # Selects a new filename to preserve previously generated workbooks.
    for index in range(1, 1000):
        filename = "student_graduation_plan.xlsx" if index == 1 else f"student_graduation_plan_{index}.xlsx"
        target = output_dir / filename
        try:
            with target.open("xb") as workbook_file:
                workbook_file.write(xlsx_bytes)
            return target
        except FileExistsError:
            continue
    raise FileExistsError("The output folder contains too many previously generated plans.")


def open_generated_workbook(file_path):
    # Opens the completed workbook in the default Windows spreadsheet application.
    file_path = Path(file_path).resolve()
    if not file_path.is_file():
        raise FileNotFoundError(f"Excel plan not found: {file_path}")
    if os.name != "nt":
        raise OSError("Automatic opening is supported on Windows. Open the file manually on this platform.")
    os.startfile(str(file_path))


def print_generated_workbook(file_path):
    # Sends the generated workbook to the default Windows print action.
    file_path = Path(file_path).resolve()
    if not file_path.is_file():
        raise FileNotFoundError(f"Excel plan not found: {file_path}")
    if os.name != "nt":
        raise OSError("Automatic printing is supported on Windows only.")
    os.startfile(str(file_path), "print")
