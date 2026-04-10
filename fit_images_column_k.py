from __future__ import annotations

import argparse
import os
import queue
import shutil
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable

DEFAULT_COLUMN_LABEL = "K"
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}

UI_COLORS = {
    "bg": "#0b0b0d",
    "panel": "#141116",
    "panel_alt": "#120b0d",
    "line": "#3a1a20",
    "text": "#f3f3f4",
    "text_muted": "#b6adb2",
    "accent": "#cf2138",
    "accent_hover": "#ea2f48",
    "accent_dark": "#8f1423",
    "log_bg": "#060607",
    "button_bg": "#22171b",
    "button_hover": "#2f1d22",
    "button_border": "#4f2a32",
}

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
class DependencyStatus:
    installed: list[str]
    missing: list[str]


@dataclass
class WorkbookTask:
    workbook_path: Path
    target_column_label: str
    target_column_index: int


@dataclass
class FileRow:
    workbook_path: Path
    row_frame: tk.Frame
    index_label: tk.Label
    path_label: tk.Label
    column_var: tk.StringVar
    column_combo: ttk.Combobox
    open_button: tk.Button
    remove_button: tk.Button


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


def column_index_to_label(index: int) -> str:
    if index <= 0:
        raise ValueError("Column index must be positive")

    value = index
    chars: list[str] = []
    while value > 0:
        value, rem = divmod(value - 1, 26)
        chars.append(chr(ord("A") + rem))

    return "".join(reversed(chars))


def build_column_choices(max_index: int = 702) -> list[str]:
    return [column_index_to_label(i) for i in range(1, max_index + 1)]


def is_excel_candidate(path: Path) -> bool:
    return (
        path.suffix.lower() in SUPPORTED_EXTENSIONS
        and not path.name.startswith("~$")
        and ".backup_" not in path.stem
    )


def find_excel_files(root: Path) -> list[Path]:
    return sorted(
        p
        for p in root.iterdir()
        if p.is_file()
        and is_excel_candidate(p)
    )


def check_dependencies() -> DependencyStatus:
    installed: list[str] = []
    missing: list[str] = []

    try:
        import win32com.client  # noqa: F401

        installed.append("pywin32")
    except ImportError:
        missing.append("pywin32 (pip install pywin32)")

    try:
        import PyInstaller  # noqa: F401

        installed.append("pyinstaller")
    except ImportError:
        missing.append("pyinstaller (pip install pyinstaller)")

    return DependencyStatus(installed=installed, missing=missing)


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


class FitImagesApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Image Fitter")
        self.geometry("1120x760")
        self.minsize(980, 680)

        icon_path = resolve_resource_path("YuRa - Copy.ico")
        if icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except Exception:
                pass

        self.file_rows: list[FileRow] = []
        self.column_choices = build_column_choices()
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.is_processing = False

        self.backup_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Chua chon file Excel")

        self._configure_theme()
        self._build_layout()
        self.after(100, self._drain_log_queue)

    def _configure_theme(self) -> None:
        self.configure(bg=UI_COLORS["bg"])

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Vertical.TScrollbar",
            troughcolor=UI_COLORS["panel_alt"],
            background=UI_COLORS["accent"],
            bordercolor=UI_COLORS["panel_alt"],
            darkcolor=UI_COLORS["accent_dark"],
            lightcolor=UI_COLORS["accent"],
            arrowcolor=UI_COLORS["text"],
        )

        style.configure(
            "Dark.TCombobox",
            fieldbackground=UI_COLORS["panel_alt"],
            background=UI_COLORS["accent"],
            foreground=UI_COLORS["text"],
            bordercolor=UI_COLORS["line"],
            lightcolor=UI_COLORS["line"],
            darkcolor=UI_COLORS["line"],
            arrowsize=14,
        )
        style.map(
            "Dark.TCombobox",
            fieldbackground=[("readonly", UI_COLORS["panel_alt"]), ("disabled", "#1d1d1f")],
            foreground=[("disabled", "#8e868a")],
            selectbackground=[("!disabled", UI_COLORS["accent"])],
            selectforeground=[("!disabled", UI_COLORS["text"])],
        )

        self.option_add("*TCombobox*Listbox.background", UI_COLORS["panel_alt"])
        self.option_add("*TCombobox*Listbox.foreground", UI_COLORS["text"])
        self.option_add("*TCombobox*Listbox.selectBackground", UI_COLORS["accent"])
        self.option_add("*TCombobox*Listbox.selectForeground", UI_COLORS["text"])

    def _build_button(
        self,
        parent,
        text: str,
        command,
        *,
        accent: bool,
        width: int,
        font: tuple[str, int, str] | tuple[str, int] = ("Segoe UI", 10),
    ) -> tk.Button:
        bg = UI_COLORS["accent"] if accent else UI_COLORS["button_bg"]
        hover = UI_COLORS["accent_hover"] if accent else UI_COLORS["button_hover"]
        border = UI_COLORS["accent_dark"] if accent else UI_COLORS["button_border"]

        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=UI_COLORS["text"],
            activebackground=hover,
            activeforeground=UI_COLORS["text"],
            relief="flat",
            bd=1,
            highlightthickness=1,
            highlightbackground=border,
            highlightcolor=border,
            padx=10,
            pady=7,
            cursor="hand2",
            width=width,
            font=font,
        )
        return button

    def _build_layout(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        root = tk.Frame(self, bg=UI_COLORS["bg"], padx=16, pady=14)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)
        root.rowconfigure(5, weight=2)

        title_label = tk.Label(
            root,
            text="Excel Image Fitter - Theme Red/Black",
            font=("Segoe UI", 18, "bold"),
            bg=UI_COLORS["bg"],
            fg=UI_COLORS["text"],
        )
        title_label.grid(row=0, column=0, sticky="w")

        subtitle_label = tk.Label(
            root,
            text="Moi file co cot rieng. Ban co the mo file tung dong, chon cot va RUN ALL.",
            font=("Segoe UI", 10),
            bg=UI_COLORS["bg"],
            fg=UI_COLORS["text_muted"],
        )
        subtitle_label.grid(row=1, column=0, sticky="w", pady=(2, 10))

        control_frame = tk.Frame(
            root,
            bg=UI_COLORS["panel"],
            highlightthickness=1,
            highlightbackground=UI_COLORS["line"],
            padx=10,
            pady=10,
        )
        control_frame.grid(row=2, column=0, sticky="ew")

        self.select_button = self._build_button(
            control_frame,
            text="+ Them file Excel",
            command=self.select_files,
            accent=True,
            width=18,
        )
        self.select_button.grid(row=0, column=0, padx=(0, 6))

        self.clear_button = self._build_button(
            control_frame,
            text="Xoa danh sach",
            command=self.clear_files,
            accent=False,
            width=14,
        )
        self.clear_button.grid(row=0, column=1, padx=(0, 6))

        self.check_deps_button = self._build_button(
            control_frame,
            text="Kiem tra thu vien",
            command=self.on_check_dependencies,
            accent=False,
            width=16,
        )
        self.check_deps_button.grid(row=0, column=2, padx=(0, 12))

        self.backup_checkbox = tk.Checkbutton(
            control_frame,
            text="Tao backup truoc khi sua",
            variable=self.backup_var,
            bg=UI_COLORS["panel"],
            fg=UI_COLORS["text"],
            activebackground=UI_COLORS["panel"],
            activeforeground=UI_COLORS["text"],
            selectcolor=UI_COLORS["panel_alt"],
            font=("Segoe UI", 10),
        )
        self.backup_checkbox.grid(row=0, column=3, padx=(0, 0), sticky="w")

        self.status_label = tk.Label(
            root,
            textvariable=self.status_var,
            bg=UI_COLORS["bg"],
            fg=UI_COLORS["text_muted"],
            font=("Segoe UI", 10),
        )
        self.status_label.grid(row=3, column=0, sticky="w", pady=(8, 6))

        files_panel = tk.Frame(
            root,
            bg=UI_COLORS["panel"],
            highlightthickness=1,
            highlightbackground=UI_COLORS["line"],
        )
        files_panel.grid(row=4, column=0, sticky="nsew", pady=(0, 10))
        files_panel.columnconfigure(0, weight=1)
        files_panel.rowconfigure(1, weight=1)

        header_row = tk.Frame(files_panel, bg=UI_COLORS["panel_alt"], padx=8, pady=7)
        header_row.grid(row=0, column=0, sticky="ew")
        header_row.columnconfigure(1, weight=1)

        tk.Label(
            header_row,
            text="#",
            bg=UI_COLORS["panel_alt"],
            fg=UI_COLORS["text_muted"],
            width=3,
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))
        tk.Label(
            header_row,
            text="Danh sach file Excel da chon",
            bg=UI_COLORS["panel_alt"],
            fg=UI_COLORS["text_muted"],
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=1, sticky="w")
        tk.Label(
            header_row,
            text="Cot",
            bg=UI_COLORS["panel_alt"],
            fg=UI_COLORS["text_muted"],
            width=8,
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=2, sticky="w", padx=(10, 0))
        tk.Label(
            header_row,
            text="Hanh dong",
            bg=UI_COLORS["panel_alt"],
            fg=UI_COLORS["text_muted"],
            width=20,
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=3, sticky="w", padx=(12, 0))

        rows_holder = tk.Frame(files_panel, bg=UI_COLORS["panel"]) 
        rows_holder.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        rows_holder.columnconfigure(0, weight=1)
        rows_holder.rowconfigure(0, weight=1)

        self.rows_canvas = tk.Canvas(
            rows_holder,
            bg=UI_COLORS["panel"],
            highlightthickness=0,
            bd=0,
            insertbackground=UI_COLORS["text"],
        )
        self.rows_canvas.grid(row=0, column=0, sticky="nsew")

        rows_scrollbar = ttk.Scrollbar(rows_holder, orient="vertical", command=self.rows_canvas.yview)
        rows_scrollbar.grid(row=0, column=1, sticky="ns")
        self.rows_canvas.configure(yscrollcommand=rows_scrollbar.set)

        self.rows_inner = tk.Frame(self.rows_canvas, bg=UI_COLORS["panel"])
        self.rows_window_id = self.rows_canvas.create_window((0, 0), window=self.rows_inner, anchor="nw")
        self.rows_inner.columnconfigure(0, weight=1)
        self.rows_inner.bind("<Configure>", self._on_rows_inner_configure)
        self.rows_canvas.bind("<Configure>", self._on_rows_canvas_configure)

        self.rows_canvas.bind("<Enter>", self._bind_rows_mousewheel)
        self.rows_canvas.bind("<Leave>", self._unbind_rows_mousewheel)

        self.empty_label = tk.Label(
            self.rows_inner,
            text="Chua co file nao. Bam '+ Them file Excel' de bat dau.",
            bg=UI_COLORS["panel"],
            fg=UI_COLORS["text_muted"],
            font=("Segoe UI", 11),
            pady=22,
        )
        self.empty_label.grid(row=0, column=0, sticky="ew")

        run_panel = tk.Frame(
            root,
            bg=UI_COLORS["panel"],
            highlightthickness=1,
            highlightbackground=UI_COLORS["line"],
            padx=10,
            pady=10,
        )
        run_panel.grid(row=5, column=0, sticky="ew", pady=(0, 10))

        self.run_button = self._build_button(
            run_panel,
            text="RUN ALL",
            command=self.start_processing,
            accent=True,
            width=18,
            font=("Segoe UI", 12, "bold"),
        )
        self.run_button.grid(row=0, column=0, sticky="w")

        log_panel = tk.Frame(
            root,
            bg=UI_COLORS["panel"],
            highlightthickness=1,
            highlightbackground=UI_COLORS["line"],
            padx=8,
            pady=8,
        )
        log_panel.grid(row=6, column=0, sticky="nsew")
        root.rowconfigure(6, weight=2)
        log_panel.columnconfigure(0, weight=1)
        log_panel.rowconfigure(1, weight=1)

        tk.Label(
            log_panel,
            text="LOG",
            bg=UI_COLORS["panel"],
            fg=UI_COLORS["text"],
            font=("Segoe UI", 11, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        self.log_text = tk.Text(
            log_panel,
            height=10,
            state="disabled",
            wrap="word",
            bg=UI_COLORS["log_bg"],
            fg=UI_COLORS["text"],
            insertbackground=UI_COLORS["text"],
            relief="flat",
            bd=0,
            padx=10,
            pady=10,
            font=("Consolas", 10),
        )
        self.log_text.grid(row=1, column=0, sticky="nsew")

        log_scroll = ttk.Scrollbar(log_panel, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=1, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set)

    def _on_rows_inner_configure(self, _event=None) -> None:
        self.rows_canvas.configure(scrollregion=self.rows_canvas.bbox("all"))

    def _on_rows_canvas_configure(self, event: tk.Event) -> None:
        self.rows_canvas.itemconfigure(self.rows_window_id, width=event.width)

    def _on_rows_mousewheel(self, event: tk.Event) -> None:
        self.rows_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _bind_rows_mousewheel(self, _event=None) -> None:
        self.rows_canvas.bind_all("<MouseWheel>", self._on_rows_mousewheel)

    def _unbind_rows_mousewheel(self, _event=None) -> None:
        self.rows_canvas.unbind_all("<MouseWheel>")

    def _is_duplicate_file(self, path: Path) -> bool:
        return any(row.workbook_path == path for row in self.file_rows)

    def _add_file_row(self, workbook_path: Path) -> None:
        row_frame = tk.Frame(
            self.rows_inner,
            bg=UI_COLORS["panel"],
            highlightthickness=1,
            highlightbackground=UI_COLORS["line"],
            padx=8,
            pady=6,
        )
        row_frame.columnconfigure(1, weight=1)

        index_label = tk.Label(
            row_frame,
            text="0",
            width=3,
            anchor="w",
            bg=UI_COLORS["panel"],
            fg=UI_COLORS["text"],
            font=("Segoe UI", 10, "bold"),
        )
        index_label.grid(row=0, column=0, sticky="w", padx=(0, 8))

        path_label = tk.Label(
            row_frame,
            text=str(workbook_path),
            anchor="w",
            justify="left",
            bg=UI_COLORS["panel"],
            fg=UI_COLORS["text"],
            font=("Segoe UI", 10),
        )
        path_label.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        column_var = tk.StringVar(value=DEFAULT_COLUMN_LABEL)
        column_combo = ttk.Combobox(
            row_frame,
            textvariable=column_var,
            values=self.column_choices,
            width=7,
            style="Dark.TCombobox",
            state="normal",
        )
        column_combo.grid(row=0, column=2, sticky="w", padx=(0, 10))

        action_frame = tk.Frame(row_frame, bg=UI_COLORS["panel"])
        action_frame.grid(row=0, column=3, sticky="w")

        open_button = self._build_button(
            action_frame,
            text="Mo Excel",
            command=lambda p=workbook_path: self.open_excel_file(p),
            accent=False,
            width=10,
        )
        open_button.grid(row=0, column=0, padx=(0, 6))

        remove_button = self._build_button(
            action_frame,
            text="Xoa",
            command=lambda p=workbook_path: self.remove_file_by_path(p),
            accent=False,
            width=8,
        )
        remove_button.grid(row=0, column=1)

        file_row = FileRow(
            workbook_path=workbook_path,
            row_frame=row_frame,
            index_label=index_label,
            path_label=path_label,
            column_var=column_var,
            column_combo=column_combo,
            open_button=open_button,
            remove_button=remove_button,
        )
        self.file_rows.append(file_row)
        self._reflow_rows()

    def _reflow_rows(self) -> None:
        if not self.file_rows:
            self.empty_label.grid(row=0, column=0, sticky="ew")
            self.status_var.set("Chua chon file Excel")
            return

        self.empty_label.grid_remove()
        for index, row in enumerate(self.file_rows, start=1):
            row.index_label.configure(text=str(index))
            row.row_frame.grid(row=index - 1, column=0, sticky="ew", padx=6, pady=4)

        self.status_var.set(f"Da chon {len(self.file_rows)} file Excel")

    def select_files(self) -> None:
        raw_paths = filedialog.askopenfilenames(
            title="Chon file Excel",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if not raw_paths:
            return

        added = 0
        skipped_invalid = 0
        skipped_duplicate = 0
        for raw_path in raw_paths:
            path = Path(raw_path).expanduser().resolve()
            if self._is_duplicate_file(path):
                skipped_duplicate += 1
                continue
            if not is_excel_candidate(path):
                skipped_invalid += 1
                continue
            self._add_file_row(path)
            added += 1

        self._enqueue_log(f"Da them {added} file vao danh sach")
        if skipped_duplicate > 0:
            self._enqueue_log(f"Bo qua {skipped_duplicate} file bi trung")
        if skipped_invalid > 0:
            self._enqueue_log(f"Bo qua {skipped_invalid} file khong dung dinh dang Excel")

    def clear_files(self) -> None:
        if self.is_processing:
            return

        for row in self.file_rows:
            row.row_frame.destroy()
        self.file_rows.clear()
        self._reflow_rows()
        self._enqueue_log("Da xoa danh sach file")

    def remove_file_by_path(self, workbook_path: Path) -> None:
        if self.is_processing:
            return

        for index, row in enumerate(self.file_rows):
            if row.workbook_path == workbook_path:
                row.row_frame.destroy()
                self.file_rows.pop(index)
                self._reflow_rows()
                self._enqueue_log(f"Da xoa file: {workbook_path.name}")
                return

    def on_check_dependencies(self) -> None:
        deps = check_dependencies()

        for package_name in deps.installed:
            self._enqueue_log(f"OK: {package_name}")

        if deps.missing:
            for package_name in deps.missing:
                self._enqueue_log(f"THIEU: {package_name}")
            messagebox.showwarning("Thieu thu vien", "\n".join(deps.missing))
            return

        messagebox.showinfo("Thu vien", "Da du thu vien: pywin32, pyinstaller")

    def open_excel_file(self, workbook_path: Path) -> None:
        if not workbook_path.exists():
            messagebox.showwarning("Khong tim thay file", str(workbook_path))
            self._enqueue_log(f"Khong mo duoc file: {workbook_path}")
            return

        try:
            os.startfile(str(workbook_path))
            self._enqueue_log(f"Da mo file: {workbook_path.name}")
        except Exception as exc:
            messagebox.showerror("Loi mo file", str(exc))
            self._enqueue_log(f"Loi khi mo file {workbook_path.name}: {exc}")

    def _collect_tasks_from_rows(self) -> list[WorkbookTask]:
        tasks: list[WorkbookTask] = []

        for row in self.file_rows:
            if not row.workbook_path.exists():
                self._enqueue_log(f"{row.workbook_path.name}: bo qua vi file khong ton tai")
                continue

            try:
                column_label = normalize_column_label(row.column_var.get())
                column_index = column_label_to_index(column_label)
            except ValueError as exc:
                raise ValueError(f"{row.workbook_path.name}: {exc}") from exc

            tasks.append(
                WorkbookTask(
                    workbook_path=row.workbook_path,
                    target_column_label=column_label,
                    target_column_index=column_index,
                )
            )

        return tasks

    def start_processing(self) -> None:
        if self.is_processing:
            return

        if not self.file_rows:
            messagebox.showwarning("Chua chon file", "Hay chon it nhat 1 file Excel")
            return

        try:
            tasks = self._collect_tasks_from_rows()
        except ValueError as exc:
            messagebox.showerror("Cot khong hop le", str(exc))
            return

        if not tasks:
            messagebox.showwarning("Khong co file hop le", "Danh sach file khong con ton tai")
            return

        self.is_processing = True
        self._set_controls_enabled(False)
        self._enqueue_log("--- Bat dau xu ly ---")

        worker = threading.Thread(
            target=self._process_worker,
            args=(tasks, self.backup_var.get()),
            daemon=True,
        )
        worker.start()

    def _process_worker(self, tasks: list[WorkbookTask], make_backup: bool) -> None:
        results: list[WorkbookResult] = []
        error_text = ""

        try:
            results = process_workbooks(
                tasks=tasks,
                make_backup=make_backup,
                logger=self._enqueue_log,
            )
        except Exception as exc:
            error_text = str(exc)

        self.after(0, lambda: self._on_processing_done(results, error_text))

    def _on_processing_done(self, results: list[WorkbookResult], error_text: str) -> None:
        if error_text:
            self._enqueue_log(f"LOI: {error_text}")
            messagebox.showerror("Loi xu ly", error_text)

        total_resized, total_pictures, total_errors = summarize_results(results)
        self._enqueue_log("---")
        self._enqueue_log(f"Tong so anh da can chinh: {total_resized}")
        self._enqueue_log(f"Tong so shape anh tim thay: {total_pictures}")
        self._enqueue_log(f"Tong so loi bo qua: {total_errors}")

        self.status_var.set(
            f"Hoan tat: {len(results)} file, resize={total_resized}, errors={total_errors}"
        )
        self.is_processing = False
        self._set_controls_enabled(True)

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        self.select_button.configure(state=state)
        self.clear_button.configure(state=state)
        self.check_deps_button.configure(state=state)
        self.backup_checkbox.configure(state=state)
        self.run_button.configure(state=state)

        combo_state = "normal" if enabled else "disabled"
        for row in self.file_rows:
            row.column_combo.configure(state=combo_state)
            row.open_button.configure(state=state)
            row.remove_button.configure(state=state)

    def _enqueue_log(self, message: str) -> None:
        self.log_queue.put(message)

    def _drain_log_queue(self) -> None:
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self._append_log(message)

        self.after(120, self._drain_log_queue)

    def _append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Resize images in selected column to exactly fit each target cell."
    )
    parser.add_argument(
        "files",
        nargs="*",
        help="Optional Excel file paths for CLI mode.",
    )
    parser.add_argument(
        "--column",
        default=DEFAULT_COLUMN_LABEL,
        help="Excel column label for CLI mode (default: K)",
    )
    parser.add_argument(
        "--cli",
        action="store_true",
        help="Run in CLI mode instead of the UI.",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Do not create backup files before editing.",
    )
    return parser.parse_args()


def run_cli(args: argparse.Namespace) -> None:
    if args.files:
        excel_files = [Path(raw_path).expanduser().resolve() for raw_path in args.files]
    else:
        root = Path(__file__).resolve().parent
        excel_files = find_excel_files(root)

    if not excel_files:
        print("Khong tim thay file Excel nao de xu ly.")
        return

    try:
        column_label = normalize_column_label(args.column)
        column_index = column_label_to_index(column_label)
    except ValueError as exc:
        print(f"Cot khong hop le: {exc}")
        return

    tasks = [
        WorkbookTask(
            workbook_path=path,
            target_column_label=column_label,
            target_column_index=column_index,
        )
        for path in excel_files
    ]

    results = process_workbooks(
        tasks=tasks,
        make_backup=not args.no_backup,
        logger=print,
    )

    total_resized, total_pictures, total_errors = summarize_results(results)
    print("---")
    print(f"Tong so anh da can chinh cot {column_label}: {total_resized}")
    print(f"Tong so shape anh tim thay: {total_pictures}")
    print(f"Tong so loi bo qua: {total_errors}")


def run_ui() -> None:
    app = FitImagesApp()
    app.mainloop()


def main() -> None:
    args = parse_args()
    if args.cli or args.files:
        run_cli(args)
        return

    run_ui()


if __name__ == "__main__":
    main()
