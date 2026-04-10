from __future__ import annotations

import os
import queue
import shutil
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import time
from typing import Callable
import tkinter as tk
from tkinter import filedialog
import webbrowser
import webview

from flask import Flask, render_template, request, jsonify

DEFAULT_COLUMN_LABEL = "K"
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}

# Office constants used through COM
MSO_SHAPE_PICTURE = 13
MSO_SHAPE_LINKED_PICTURE = 11
MSO_FALSE = 0
XL_MOVE_AND_SIZE = 1
RPC_E_CALL_REJECTED = -2147418111


def resolve_resource_path(name: str) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS")) / name
    return Path(__file__).resolve().with_name(name)


@dataclass
class WorkbookResult:
    file_name: str
    target_column: str
    backup_path: str
    resized_images: int
    pictures_found: int
    errors: int


@dataclass
class WorkbookTask:
    workbook_path: Path
    target_column_label: str
    target_column_index: int


def normalize_column_label(raw_value: str) -> str:
    label = raw_value.strip().upper()
    if not label:
        raise ValueError("Cot khong duoc de trong")
    if not label.isalpha():
        raise ValueError("Cot chi duoc gom cac ky tu A-Z")
    if len(label) > 3:
        raise ValueError("Cot qua dai. Vi du hop le: K, AA, XFD")
    return label


def column_label_to_index(raw_value: str) -> int:
    label = normalize_column_label(raw_value)
    value = 0
    for char in label:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def is_excel_candidate(path: Path) -> bool:
    return (
        path.suffix.lower() in SUPPORTED_EXTENSIONS
        and not path.name.startswith("~$")
        and ".backup_" not in path.stem
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


def fit_images_in_column(
    excel_app,
    workbook_path: Path,
    target_column_index: int,
    target_column_label: str,
    make_backup: bool,
) -> WorkbookResult:
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
                    if int(top_left_cell.Column) != target_column_index:
                        continue

                    target_cell = com_retry(
                        lambda: worksheet.Cells(int(top_left_cell.Row), target_column_index)
                    )

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
        target_column=target_column_label,
        backup_path=backup_path,
        resized_images=resized_images,
        pictures_found=pictures_found,
        errors=errors,
    )


def process_workbooks(
    tasks: list[WorkbookTask],
    make_backup: bool,
    logger: Callable[[str], None],
) -> list[WorkbookResult]:
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("Thieu thu vien pywin32. Cai dat bang lenh: pip install pywin32") from exc

    excel_app = com_retry(lambda: win32.DispatchEx("Excel.Application"))

    def set_app_property(name: str, value) -> None:
        try:
            com_retry(lambda: setattr(excel_app, name, value))
        except Exception:
            pass

    set_app_property("Visible", False)
    set_app_property("DisplayAlerts", False)
    set_app_property("ScreenUpdating", False)
    set_app_property("EnableEvents", False)

    results: list[WorkbookResult] = []
    try:
        for task in tasks:
            excel_file = task.workbook_path

            if not excel_file.exists():
                logger(f"[{task.target_column_label}] {excel_file.name}: bo qua vi file khong ton tai")
                continue

            if not is_excel_candidate(excel_file):
                logger(f"[{task.target_column_label}] {excel_file.name}: bo qua vi khong phai file Excel hop le")
                continue

            try:
                result = fit_images_in_column(
                    excel_app=excel_app,
                    workbook_path=excel_file,
                    target_column_index=task.target_column_index,
                    target_column_label=task.target_column_label,
                    make_backup=make_backup,
                )
                results.append(result)
                logger(
                    f"[{result.target_column}] {result.file_name}: resize={result.resized_images}, "
                    f"pictures={result.pictures_found}, errors={result.errors}"
                )
                logger(f"  backup: {result.backup_path}")
            except Exception as exc:
                logger(f"[{task.target_column_label}] {excel_file.name}: loi -> {exc}")
    finally:
        com_retry(lambda: excel_app.Quit())

    return results

def summarize_results(results: list[WorkbookResult]) -> tuple[int, int, int]:
    total_resized = sum(r.resized_images for r in results)
    total_pictures = sum(r.pictures_found for r in results)
    total_errors = sum(r.errors for r in results)
    return total_resized, total_pictures, total_errors


# --- FLASK APP AND ROUTES ---

if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    static_folder = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
else:
    app = Flask(__name__)
log_queue = queue.Queue()

def web_logger(message: str):
    log_queue.put(message)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/select_files', methods=['POST'])
def api_select_files():
    # Use standard tkinter dialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_paths = filedialog.askopenfilenames(
        title="Chon file Excel",
        filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xls"), ("All files", "*.*")],
    )
    root.destroy()
    return jsonify({"files": list(file_paths)})

@app.route('/api/run', methods=['POST'])
def api_run():
    data = request.json
    make_backup = data.get('make_backup', True)
    raw_tasks = data.get('tasks', [])
    
    tasks = []
    for item in raw_tasks:
        try:
            path = Path(item['path'])
            col_label = normalize_column_label(item['column'])
            col_idx = column_label_to_index(col_label)
            tasks.append(WorkbookTask(path, col_label, col_idx))
        except Exception as e:
            web_logger(f"Invalid column format for {item['path']}: {e}")

    if not tasks:
        return jsonify({"status": "error", "message": "No valid tasks provided."})

    def run_worker():
        try:
            web_logger("Processing started...")
            results = process_workbooks(tasks, make_backup, web_logger)
            tr, tp, te = summarize_results(results)
            web_logger(f"DONE. Resized: {tr}, Pictures: {tp}, Errors: {te}")
        except Exception as e:
            web_logger(f"Critical error: {e}")

    thread = threading.Thread(target=run_worker, daemon=True)
    thread.start()
    thread.join() # Wait or let it run async (we wait here for simplicity so frontend poll gets it)

    # Note: process_workbooks runs synchronously in thread.join() to maintain response state.
    # In a full async real-world app, we would return immediately and let the frontend poll status.
    # For now this is fine since it's local.
    
    return jsonify({
        "status": "success", 
        "total_resized": 0, # Since we get logs via queue
        "total_pictures": 0,
        "total_errors": 0
    })

@app.route('/api/logs', methods=['GET'])
def api_logs():
    logs = []
    while not log_queue.empty():
        try:
            logs.append(log_queue.get_nowait())
        except queue.Empty:
            break
    return jsonify({"logs": logs})

@app.route('/api/open_file', methods=['POST'])
def api_open_file():
    data = request.json
    file_path = data.get('path')
    if not file_path or not os.path.exists(file_path):
        return jsonify({"status": "error", "message": "File path invalid or missing."})

    def open_worker():
        try:
            import win32com.client as win32
        except ImportError:
            web_logger("[ERROR] Missing pywin32 to control Excel.")
            return

        try:
            excel_app = win32.Dispatch("Excel.Application")
            
            # Save and close any currently open workbooks
            wb_count = excel_app.Workbooks.Count
            if wb_count > 0:
                web_logger(f"[SYS] Saving and closing {wb_count} open workbook(s)...")
                for i in range(wb_count, 0, -1):
                    try:
                        wb = excel_app.Workbooks(i)
                        wb_name = wb.Name
                        wb.Save()
                        wb.Close()
                        web_logger(f"Closed: {wb_name}")
                    except Exception as e:
                        web_logger(f"[ERROR] Failed closing internal wb: {e}")
            
            web_logger(f"Launching workbook: {os.path.basename(file_path)}...")
            excel_app.Workbooks.Open(file_path)
            excel_app.Visible = True
            web_logger("[SYS] Workbook is now open & visible.")

        except Exception as e:
            web_logger(f"[ERROR] COM Exception: {e}")

    threading.Thread(target=open_worker, daemon=True).start()
    return jsonify({"status": "success", "message": "Dispatched open command."})

if __name__ == "__main__":
    # Start the native desktop window wrapping the Flask app
    window = webview.create_window(
        'NEXUS COMMAND PRO', 
        app, 
        width=1280, 
        height=800,
        background_color='#010405'
    )
    webview.start()
