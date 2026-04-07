from __future__ import annotations

import argparse
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import time

TARGET_COLUMN_INDEX = 11  # Column K
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}

# Office constants used through COM
MSO_SHAPE_PICTURE = 13
MSO_SHAPE_LINKED_PICTURE = 11
MSO_FALSE = 0
XL_MOVE_AND_SIZE = 1
RPC_E_CALL_REJECTED = -2147418111


@dataclass
class WorkbookResult:
    file_name: str
    backup_path: str
    resized_images: int
    pictures_found: int
    errors: int


def find_excel_files(root: Path) -> list[Path]:
    return sorted(
        p
        for p in root.iterdir()
        if p.is_file()
        and p.suffix.lower() in SUPPORTED_EXTENSIONS
        and not p.name.startswith("~$")
        and ".backup_" not in p.stem
    )


def create_backup_file(workbook_path: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = workbook_path.with_name(f"{workbook_path.stem}.backup_{timestamp}{workbook_path.suffix}")
    counter = 1
    while candidate.exists():
        candidate = workbook_path.with_name(
            f"{workbook_path.stem}.backup_{timestamp}_{counter}{workbook_path.suffix}"
        )
        counter += 1

    shutil.copy2(workbook_path, candidate)
    return candidate


def is_picture_shape(shape) -> bool:
    try:
        shape_type = int(shape.Type)
    except Exception:
        return False

    return shape_type in (MSO_SHAPE_PICTURE, MSO_SHAPE_LINKED_PICTURE)


def com_retry(action, attempts: int = 30, delay_seconds: float = 0.2):
    """Retry transient Excel COM busy errors."""
    last_error = None
    for _ in range(attempts):
        try:
            return action()
        except Exception as exc:
            hresult = exc.args[0] if getattr(exc, "args", None) else None
            if hresult == RPC_E_CALL_REJECTED:
                last_error = exc
                time.sleep(delay_seconds)
                continue
            raise

    if last_error is not None:
        raise last_error

    raise RuntimeError("COM action failed without exception details")


def fit_images_in_column_k(excel_app, workbook_path: Path, make_backup: bool) -> WorkbookResult:
    backup_path = "(skip backup)"
    if make_backup:
        backup_path = str(create_backup_file(workbook_path))

    workbook = None
    resized_images = 0
    pictures_found = 0
    errors = 0

    try:
        workbook = com_retry(
            lambda: excel_app.Workbooks.Open(str(workbook_path), UpdateLinks=0, ReadOnly=False)
        )

        for worksheet in com_retry(lambda: workbook.Worksheets):
            shape_count = int(com_retry(lambda: worksheet.Shapes.Count))
            for index in range(1, shape_count + 1):
                shape = com_retry(lambda: worksheet.Shapes.Item(index))
                if not is_picture_shape(shape):
                    continue

                pictures_found += 1

                try:
                    top_left_cell = com_retry(lambda: shape.TopLeftCell)
                    if int(top_left_cell.Column) != TARGET_COLUMN_INDEX:
                        continue

                    target_cell = com_retry(lambda: worksheet.Cells(int(top_left_cell.Row), TARGET_COLUMN_INDEX))

                    # Resize directly by cell dimensions, then lock movement with the cell.
                    com_retry(lambda: setattr(shape, "LockAspectRatio", MSO_FALSE))
                    com_retry(lambda: setattr(shape, "Placement", XL_MOVE_AND_SIZE))
                    com_retry(lambda: setattr(shape, "Left", float(target_cell.Left)))
                    com_retry(lambda: setattr(shape, "Top", float(target_cell.Top)))
                    com_retry(lambda: setattr(shape, "Width", float(target_cell.Width)))
                    com_retry(lambda: setattr(shape, "Height", float(target_cell.Height)))
                    resized_images += 1
                except Exception:
                    errors += 1

        com_retry(lambda: workbook.Save())
    finally:
        if workbook is not None:
            com_retry(lambda: workbook.Close(SaveChanges=False))

    return WorkbookResult(
        file_name=workbook_path.name,
        backup_path=backup_path,
        resized_images=resized_images,
        pictures_found=pictures_found,
        errors=errors,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Find images in column K and resize each image to fit its cell exactly."
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Do not create backup files before editing.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    root = Path(__file__).resolve().parent
    excel_files = find_excel_files(root)

    if not excel_files:
        print("Khong tim thay file Excel nao de xu ly.")
        return

    try:
        import win32com.client as win32
    except ImportError:
        print("Thieu thu vien pywin32. Cai dat bang lenh: pip install pywin32")
        return

    excel_app = com_retry(lambda: win32.DispatchEx("Excel.Application"))
    excel_app.Visible = False
    excel_app.DisplayAlerts = False
    excel_app.ScreenUpdating = False
    excel_app.EnableEvents = False

    results: list[WorkbookResult] = []
    try:
        for excel_file in excel_files:
            try:
                result = fit_images_in_column_k(
                    excel_app=excel_app,
                    workbook_path=excel_file,
                    make_backup=not args.no_backup,
                )
                results.append(result)
                print(
                    f"{result.file_name}: resize={result.resized_images}, "
                    f"pictures={result.pictures_found}, errors={result.errors}"
                )
                print(f"  backup: {result.backup_path}")
            except Exception as exc:
                print(f"{excel_file.name}: loi -> {exc}")
    finally:
        com_retry(lambda: excel_app.Quit())

    total_resized = sum(r.resized_images for r in results)
    total_pictures = sum(r.pictures_found for r in results)
    total_errors = sum(r.errors for r in results)

    print("---")
    print(f"Tong so anh da can chinh cot K: {total_resized}")
    print(f"Tong so shape anh tim thay: {total_pictures}")
    print(f"Tong so loi bo qua: {total_errors}")


if __name__ == "__main__":
    main()
