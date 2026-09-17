"""Excel 管理級數整理工具。

此程式需在 Windows 且已安裝 Microsoft Excel 的環境執行，透過 pywin32 讀取
Excel 原生 COM 屬性 Interior.ColorIndex，以保留格式、公式、巨集與工作表設定。
"""

from __future__ import annotations

import os
import queue
import re
import tempfile
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


DEFAULT_RANGE = "M1:R2000"
MANAGEMENT_HEADER = "管理級數"
LEVEL_HEADERS = ["第一級", "第二級", "第三級", "第四級"]
COLOR_TO_LEVEL = {-4142: "第一級", 34: "第二級", 6: "第三級", 3: "第四級"}
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
SAVE_FORMATS = {".xlsx": 51, ".xlsm": 52, ".xls": 56}

HELP_SECTIONS = [
    ("一、程式用途", """本工具用於整理 Excel 指定範圍中的儲存格資料，依照 Excel 原生的 Interior.ColorIndex 判斷管理級數，並將符合條件的欄位表頭整理到「第一級」、「第二級」、「第三級」、「第四級」對應欄位中。"""),
    ("二、支援的 Excel 格式", """支援：
• .xlsx
• .xlsm
• .xls

建議優先使用 .xlsx。
如果使用 .xlsm，程式會保留原有 VBA 巨集。"""),
    ("三、匯入 Excel 前注意事項", """1. Excel 第一列必須為表頭。
2. 工作表第一列必須存在「管理級數」表頭。
3. 表頭名稱必須完全一致，不可多空格、少字或使用其他名稱。
   正確：管理級數
   錯誤：管理 級數、管理級、管理級數　
4. 欲判斷的資料欄位也必須在第一列有表頭名稱。
5. 欲處理的資料必須位於使用者指定的資料範圍內。例如 M1:R2000 只會判斷 M 到 R 欄、1 到 2000 列的內容。
6. 指定範圍以第一列作為表頭，實際資料從第二列開始判斷。
7. 空白儲存格不會進行管理級數判斷。
8. 儲存格即使有底色，只要沒有任何資料，仍視為空白，不會列入結果。
9. 建議處理前先將 Excel 檔案關閉，避免檔案被鎖定而無法儲存。
10. 建議先保留原始 Excel 備份。"""),
    ("四、ColorIndex 判斷規則", """ColorIndex -4142 → 第一級
ColorIndex 34    → 第二級
ColorIndex 6     → 第三級
ColorIndex 3     → 第四級

其他 ColorIndex：不處理。

本程式判斷的是 Excel 原生 Interior.ColorIndex，不是單純依照畫面看到的 RGB 顏色判斷。"""),
    ("五、整理結果規則", """當某一個儲存格符合管理級數條件時，程式會取得該欄位第一列的表頭名稱，並寫入相同資料列的管理級數欄位。

例如：F1 = 噪音作業、F4 有資料，且 F4 的 ColorIndex = -4142，則該列「第一級」欄位會寫入：
噪音作業；

如果同一列有多個第一級項目，例如「噪音作業；」與「粉塵作業；」，最後會整理成：
噪音作業；粉塵作業；

排列順序依 Excel 欄位由左至右。"""),
    ("六、四個級數欄位", """程式會在「管理級數」後方建立：
第一級
第二級
第三級
第四級

如果這四個欄位已經存在，不會重複新增。
每次重新執行時，程式會先清除這四個欄位原本的整理結果，再重新計算，因此不會重複累加舊資料。"""),
    ("七、資料範圍說明", """資料範圍可自行輸入。例如 M1:R2000 代表：
起始欄：M
結束欄：R
起始列：1
結束列：2000

程式只會判斷此範圍內的儲存格，範圍外的資料不會被判斷。如果輸入格式錯誤，程式會停止處理並顯示錯誤訊息。"""),
    ("八、輸出檔案", """程式原則上使用另存新檔，不直接修改原始檔案。
請確認輸出位置有寫入權限。處理完成後，請確認輸出檔案是否正常開啟。"""),
    ("九、常見錯誤", """1. 問題：找不到「管理級數」
   原因：第一列表頭不存在完全相同的「管理級數」。

2. 問題：某些資料沒有被分類
   可能原因：
   • 儲存格為空白
   • ColorIndex 不是指定的四種
   • 資料不在指定範圍內
   • 表頭為空白
   • Excel 使用的是條件格式，看起來有顏色但原始 ColorIndex 不符合

3. 問題：Excel 無法儲存
   可能原因：
   • Excel 檔案目前已開啟
   • 檔案被其他程式鎖定
   • 輸出資料夾沒有寫入權限
   • 檔案名稱或路徑異常

4. 問題：處理結果與預期不同
   建議先確認：
   • 指定的資料範圍
   • 第一列表頭
   • 儲存格實際 ColorIndex
   • 儲存格是否真的有資料"""),
    ("十、建議操作流程", """1. 準備 Excel 檔案
2. 確認第一列為表頭
3. 確認有「管理級數」
4. 關閉 Excel 檔案
5. 開啟本工具
6. 選擇 Excel 檔案
7. 選擇工作表
8. 輸入資料範圍
9. 選擇輸出位置
10. 按下開始整理
11. 等待處理完成
12. 開啟輸出 Excel 確認結果"""),
]


class ExcelProcessError(Exception):
    """可顯示給使用者的 Excel 處理錯誤。"""


@dataclass(frozen=True)
class ExcelRange:
    """使用者輸入的 Excel 範圍。"""

    start_col: int
    start_row: int
    end_col: int
    end_row: int

    @property
    def row_count(self) -> int:
        return self.end_row - self.start_row + 1

    @property
    def data_row_count(self) -> int:
        return max(0, self.end_row - self.start_row)


@dataclass
class ProcessStats:
    """整理完成後回傳給 GUI 的統計資訊。"""

    worksheet_name: str = ""
    actual_range: str = ""
    scanned_rows: int = 0
    non_empty_cells: int = 0
    level_counts: dict[str, int] = field(default_factory=lambda: {name: 0 for name in LEVEL_HEADERS})
    ignored_color_cells: int = 0
    output_path: str = ""

    def to_message(self) -> str:
        return (
            f"實際處理的工作表名稱：{self.worksheet_name}\n"
            f"實際處理的範圍：{self.actual_range}\n"
            f"掃描的資料列數：{self.scanned_rows}\n"
            f"掃描的非空白儲存格數：{self.non_empty_cells}\n"
            f"第一級寫入項目數：{self.level_counts['第一級']}\n"
            f"第二級寫入項目數：{self.level_counts['第二級']}\n"
            f"第三級寫入項目數：{self.level_counts['第三級']}\n"
            f"第四級寫入項目數：{self.level_counts['第四級']}\n"
            f"忽略的其他 ColorIndex 儲存格數：{self.ignored_color_cells}\n"
            f"輸出檔案路徑：{self.output_path}"
        )


def column_letter_to_index(column_letters: str) -> int:
    """將欄名（例如 M、AA）轉為 1-based 欄號。"""
    value = 0
    for char in column_letters.upper():
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def column_index_to_letter(column_index: int) -> str:
    """將 1-based 欄號轉為 Excel 欄名。"""
    letters = ""
    while column_index:
        column_index, remainder = divmod(column_index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def validate_range(range_text: str) -> ExcelRange:
    """驗證並解析 A1:B2 形式的資料範圍。"""
    match = re.fullmatch(r"\s*([A-Za-z]{1,3})(\d+)\s*:\s*([A-Za-z]{1,3})(\d+)\s*", range_text or "")
    if not match:
        raise ExcelProcessError("資料範圍格式錯誤，請使用類似 M1:R2000 的格式。")

    start_col = column_letter_to_index(match.group(1))
    start_row = int(match.group(2))
    end_col = column_letter_to_index(match.group(3))
    end_row = int(match.group(4))

    if start_col > end_col or start_row > end_row:
        raise ExcelProcessError("資料範圍格式錯誤，請使用類似 M1:R2000 的格式。")
    if start_col < 1 or end_col > EXCEL_MAX_COLUMNS or start_row < 1 or end_row > EXCEL_MAX_ROWS:
        raise ExcelProcessError("指定範圍超出 Excel 可使用範圍。")
    if start_row != 1:
        raise ExcelProcessError("範圍起始列必須為第 1 列，因為第一列為來源表頭。")
    return ExcelRange(start_col, start_row, end_col, end_row)


def format_range(excel_range: ExcelRange) -> str:
    """將 ExcelRange 轉回 A1:B2 字串。"""
    return f"{column_index_to_letter(excel_range.start_col)}{excel_range.start_row}:{column_index_to_letter(excel_range.end_col)}{excel_range.end_row}"


def is_blank(value: Any) -> bool:
    """判斷 None、空字串、只有空白字元的字串是否為空白。"""
    return value is None or (isinstance(value, str) and value.strip() == "")


def normalize_color_index(color_index: Any) -> int | None:
    """將 COM 回傳的 ColorIndex 轉為整數，無法轉換時回傳 None。"""
    try:
        return int(color_index)
    except (TypeError, ValueError):
        return None


def default_output_path(input_path: str) -> str:
    """依原始副檔名產生預設輸出路徑，xlsm 會維持 xlsm。"""
    path = Path(input_path)
    extension = path.suffix.lower()
    output_extension = extension if extension in {".xlsm", ".xls"} else ".xlsx"
    return str(path.with_name(f"{path.stem}_整理完成{output_extension}"))


class ExcelProcessor:
    """封裝所有 Excel COM 操作，避免 GUI 與資料處理邏輯混在一起。"""

    def __init__(self, progress_callback=None, log_callback=None) -> None:
        self.progress_callback = progress_callback
        self.log_callback = log_callback
        self.excel = None
        self.workbook = None

    def log(self, message: str) -> None:
        if self.log_callback:
            self.log_callback(message)

    def update_progress(self, value: int, maximum: int) -> None:
        if self.progress_callback:
            self.progress_callback(value, maximum)

    def load_worksheet_names(self, file_path: str) -> list[str]:
        """以獨立 Excel 執行個體讀取工作表名稱。"""
        self._validate_input_file(file_path)
        self._open_excel(file_path, read_only=True)
        try:
            return [sheet.Name for sheet in self.workbook.Worksheets]
        finally:
            self.release_excel_objects(save_changes=False)

    def process_excel_data(self, file_path: str, worksheet_name: str, range_text: str, output_path: str) -> ProcessStats:
        """主要整理流程。"""
        original_range = validate_range(range_text)
        self._validate_paths(file_path, output_path)
        self._open_excel(file_path, read_only=False)
        try:
            worksheet = self._get_worksheet(worksheet_name)
            self._ensure_writable(output_path)
            source_headers = self._capture_source_headers(worksheet, original_range)
            management_col = self.find_management_level_column(worksheet)
            level_columns, inserted_count = self.create_level_columns(worksheet, management_col)
            adjusted_range, adjusted_headers = self._adjust_source_range(original_range, source_headers, management_col, inserted_count)
            self._validate_adjusted_headers(adjusted_headers)
            self._clear_old_results(worksheet, level_columns, adjusted_range.end_row)
            stats = self._scan_and_write(worksheet, worksheet_name, adjusted_range, adjusted_headers, level_columns)
            stats.output_path = output_path
            self.save_output_file(output_path)
            return stats
        finally:
            self.release_excel_objects(save_changes=False)

    def find_management_level_column(self, worksheet) -> int:
        """在第一列尋找「管理級數」表頭。"""
        last_col = worksheet.Cells(1, worksheet.Columns.Count).End(-4159).Column  # xlToLeft
        for col in range(1, last_col + 1):
            if str(worksheet.Cells(1, col).Value or "").strip() == MANAGEMENT_HEADER:
                return col
        raise ExcelProcessError("找不到表頭「管理級數」，請確認 Excel 第一列內容。")

    def create_level_columns(self, worksheet, management_col: int) -> tuple[dict[str, int], int]:
        """建立或重用第一級至第四級欄位。"""
        existing = self._find_level_columns_after_management(worksheet, management_col)
        if all(header in existing for header in LEVEL_HEADERS):
            self.log("已存在第一級至第四級表頭，將直接重用既有欄位。")
            return {header: existing[header] for header in LEVEL_HEADERS}, 0

        insert_at = management_col + 1
        worksheet.Columns(f"{column_index_to_letter(insert_at)}:{column_index_to_letter(insert_at + 3)}").Insert(Shift=1)
        for offset, header in enumerate(LEVEL_HEADERS):
            target_col = insert_at + offset
            worksheet.Cells(1, management_col).Copy()
            worksheet.Cells(1, target_col).PasteSpecial(Paste=-4122)  # xlPasteFormats
            worksheet.Cells(1, target_col).Value = header
            worksheet.Columns(management_col).Copy()
            worksheet.Columns(target_col).PasteSpecial(Paste=-4122)
        worksheet.Application.CutCopyMode = False
        self.log("已在管理級數後方新增第一級至第四級欄位。")
        return {header: insert_at + index for index, header in enumerate(LEVEL_HEADERS)}, 4

    def get_cell_color_index(self, cell) -> int | None:
        """取得 Excel 原生 Interior.ColorIndex。"""
        return normalize_color_index(cell.Interior.ColorIndex)

    def save_output_file(self, output_path: str) -> None:
        """依副檔名另存新檔，保留 xlsm 巨集格式。"""
        extension = Path(output_path).suffix.lower()
        file_format = SAVE_FORMATS.get(extension, 51)
        self.workbook.SaveAs(os.path.abspath(output_path), FileFormat=file_format)

    def release_excel_objects(self, save_changes: bool = False) -> None:
        """關閉本程式建立的 Workbook 與 Excel Application。"""
        if self.workbook is not None:
            self.workbook.Close(SaveChanges=save_changes)
            self.workbook = None
        if self.excel is not None:
            self.excel.Quit()
            self.excel = None

    def _open_excel(self, file_path: str, read_only: bool) -> None:
        import win32com.client

        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.workbook = self.excel.Workbooks.Open(os.path.abspath(file_path), ReadOnly=read_only)

    def _validate_input_file(self, file_path: str) -> None:
        if not file_path or not Path(file_path).exists():
            raise ExcelProcessError("Excel 檔案不存在，請重新選擇檔案。")
        if Path(file_path).suffix.lower() not in SAVE_FORMATS:
            raise ExcelProcessError("僅支援 .xlsx、.xlsm、.xls 檔案。")

    def _validate_paths(self, file_path: str, output_path: str) -> None:
        self._validate_input_file(file_path)
        if not output_path:
            raise ExcelProcessError("請選擇輸出檔案路徑。")
        output_parent = Path(output_path).expanduser().resolve().parent
        if not output_parent.exists():
            raise ExcelProcessError("輸出路徑的資料夾不存在。")
        if Path(output_path).suffix.lower() not in SAVE_FORMATS:
            raise ExcelProcessError("輸出檔案副檔名需為 .xlsx、.xlsm 或 .xls。")

    def _ensure_writable(self, output_path: str) -> None:
        """以暫存檔測試輸出目錄寫入權限。"""
        directory = Path(output_path).resolve().parent
        try:
            with tempfile.NamedTemporaryFile(dir=directory, delete=True):
                pass
        except OSError as exc:
            raise ExcelProcessError(f"輸出路徑無法寫入：{exc}") from exc

    def _get_worksheet(self, worksheet_name: str):
        for sheet in self.workbook.Worksheets:
            if sheet.Name == worksheet_name:
                return sheet
        raise ExcelProcessError("工作表不存在，請重新選擇工作表。")

    def _capture_source_headers(self, worksheet, excel_range: ExcelRange) -> list[tuple[int, str]]:
        headers = []
        for col in range(excel_range.start_col, excel_range.end_col + 1):
            value = worksheet.Cells(excel_range.start_row, col).Value
            if is_blank(value):
                raise ExcelProcessError("來源範圍的表頭不可為空白。")
            headers.append((col, str(value).strip()))
        return headers

    def _find_level_columns_after_management(self, worksheet, management_col: int) -> dict[str, int]:
        found = {}
        for offset in range(1, 5):
            col = management_col + offset
            header = str(worksheet.Cells(1, col).Value or "").strip()
            if header in LEVEL_HEADERS:
                found[header] = col
        return found

    def _adjust_source_range(self, original_range: ExcelRange, headers: list[tuple[int, str]], management_col: int, inserted_count: int) -> tuple[ExcelRange, list[tuple[int, str]]]:
        if inserted_count == 0:
            return original_range, headers
        insert_at = management_col + 1
        adjusted_headers = []
        for original_col, header in headers:
            adjusted_col = original_col + inserted_count if original_col >= insert_at else original_col
            adjusted_headers.append((adjusted_col, header))
        adjusted_start = original_range.start_col + inserted_count if original_range.start_col >= insert_at else original_range.start_col
        adjusted_end = original_range.end_col + inserted_count if original_range.end_col >= insert_at else original_range.end_col
        return ExcelRange(adjusted_start, original_range.start_row, adjusted_end, original_range.end_row), adjusted_headers

    def _validate_adjusted_headers(self, headers: list[tuple[int, str]]) -> None:
        for _col, header in headers:
            if header in LEVEL_HEADERS:
                raise ExcelProcessError("來源範圍包含第一級至第四級結果欄，請調整資料範圍後再執行。")

    def _clear_old_results(self, worksheet, level_columns: dict[str, int], end_row: int) -> None:
        for col in level_columns.values():
            if end_row >= 2:
                worksheet.Range(worksheet.Cells(2, col), worksheet.Cells(end_row, col)).ClearContents()

    def _scan_and_write(self, worksheet, worksheet_name: str, excel_range: ExcelRange, headers: list[tuple[int, str]], level_columns: dict[str, int]) -> ProcessStats:
        stats = ProcessStats(worksheet_name=worksheet_name, actual_range=format_range(excel_range), scanned_rows=excel_range.data_row_count)
        maximum = max(1, excel_range.data_row_count)
        for row_index, row in enumerate(range(2, excel_range.end_row + 1), start=1):
            row_results = {header: [] for header in LEVEL_HEADERS}
            for col, source_header in headers:
                cell = worksheet.Cells(row, col)
                if is_blank(cell.Value):
                    continue
                stats.non_empty_cells += 1
                color_index = self.get_cell_color_index(cell)
                level_header = COLOR_TO_LEVEL.get(color_index)
                if level_header is None:
                    stats.ignored_color_cells += 1
                    continue
                row_results[level_header].append(source_header)
                stats.level_counts[level_header] += 1
            for level_header, names in row_results.items():
                if names:
                    worksheet.Cells(row, level_columns[level_header]).Value = "".join(f"{name}；" for name in names)
            if row_index % 10 == 0 or row_index == maximum:
                self.update_progress(row_index, maximum)
        return stats


class ExcelOrganizerApp(ctk.CTk):
    """CustomTkinter GUI 主視窗。"""

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel 管理級數整理工具")
        self.geometry("920x720")
        self.minsize(720, 600)
        self.message_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.file_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.range_var = tk.StringVar(value=DEFAULT_RANGE)
        self.worksheet_var = tk.StringVar()
        self.help_window: ctk.CTkToplevel | None = None
        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 12))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Excel 管理級數整理工具", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(header, text="使用說明", width=100, command=self.show_help_window).grid(row=0, column=1, sticky="e")

        settings = ctk.CTkFrame(self, corner_radius=12)
        settings.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 14))
        settings.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(settings, text="整理設定", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=20, pady=(16, 10))

        ctk.CTkLabel(settings, text="Excel 檔案", anchor="w").grid(row=1, column=0, sticky="w", padx=(20, 12), pady=7)
        ctk.CTkEntry(settings, textvariable=self.file_path_var).grid(row=1, column=1, sticky="ew", pady=7)
        ctk.CTkButton(settings, text="選擇檔案", width=110, command=self.select_excel_file).grid(row=1, column=2, padx=(12, 20), pady=7)

        ctk.CTkLabel(settings, text="工作表", anchor="w").grid(row=2, column=0, sticky="w", padx=(20, 12), pady=7)
        self.sheet_combo = ctk.CTkComboBox(settings, variable=self.worksheet_var, values=[], state="readonly")
        self.sheet_combo.grid(row=2, column=1, sticky="ew", pady=7)

        ctk.CTkLabel(settings, text="資料範圍", anchor="w").grid(row=3, column=0, sticky="w", padx=(20, 12), pady=7)
        ctk.CTkEntry(settings, textvariable=self.range_var).grid(row=3, column=1, sticky="ew", pady=7)
        ctk.CTkLabel(settings, text="例：M1:R2000", text_color=("gray45", "gray65")).grid(row=3, column=2, sticky="w", padx=(12, 20), pady=7)

        ctk.CTkLabel(settings, text="輸出檔案", anchor="w").grid(row=4, column=0, sticky="w", padx=(20, 12), pady=(7, 18))
        ctk.CTkEntry(settings, textvariable=self.output_path_var).grid(row=4, column=1, sticky="ew", pady=(7, 18))
        ctk.CTkButton(settings, text="選擇輸出位置", width=110, command=self.select_output_file).grid(row=4, column=2, padx=(12, 20), pady=(7, 18))

        activity = ctk.CTkFrame(self, corner_radius=12)
        activity.grid(row=2, column=0, sticky="nsew", padx=24, pady=(0, 20))
        activity.grid_columnconfigure(0, weight=1)
        activity.grid_rowconfigure(3, weight=1)
        self.start_button = ctk.CTkButton(activity, text="開始整理", height=44, corner_radius=8, font=ctk.CTkFont(size=15, weight="bold"), command=self.start_processing)
        self.start_button.grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 12))
        self.progress = ctk.CTkProgressBar(activity, mode="determinate")
        self.progress.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 16))
        self.progress.set(0)
        ctk.CTkLabel(activity, text="執行紀錄", font=ctk.CTkFont(size=15, weight="bold")).grid(row=2, column=0, sticky="w", padx=20, pady=(0, 8))
        self.log_text = ctk.CTkTextbox(activity, wrap="word", corner_radius=8)
        self.log_text.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 18))

    def show_help_window(self) -> None:
        """開啟唯一且附屬於主視窗的 modal 使用說明視窗。"""
        if self.help_window is not None and self.help_window.winfo_exists():
            self.help_window.deiconify()
            self.help_window.lift()
            self.help_window.focus_force()
            return

        self.help_window = ctk.CTkToplevel(self)
        self.help_window.title("Excel 管理級數整理工具－使用說明")
        self.help_window.geometry("700x600")
        self.help_window.minsize(520, 400)
        self.help_window.transient(self)
        self.help_window.protocol("WM_DELETE_WINDOW", self.close_help_window)

        help_text = ctk.CTkTextbox(self.help_window, wrap="word", corner_radius=8)
        help_text.grid(row=0, column=0, sticky="nsew", padx=16, pady=(16, 8))
        for heading, body in HELP_SECTIONS:
            help_text.insert("end", f"【{heading}】\n{body}\n\n")
        help_text.configure(state="disabled")

        ctk.CTkButton(self.help_window, text="關閉", width=100, command=self.close_help_window).grid(row=1, column=0, pady=(4, 16))
        self.help_window.columnconfigure(0, weight=1)
        self.help_window.rowconfigure(0, weight=1)
        self.help_window.grab_set()
        self.help_window.lift()
        self.help_window.focus_force()

        def keep_help_window_in_front() -> None:
            if self.help_window is not None and self.help_window.winfo_exists():
                self.help_window.lift()
                self.help_window.focus_force()

        self.help_window.after(100, keep_help_window_in_front)

    def close_help_window(self) -> None:
        """釋放 modal 狀態、關閉說明視窗並清除視窗參考。"""
        if self.help_window is None:
            return
        try:
            self.help_window.grab_release()
        except (tk.TclError, RuntimeError):
            pass
        self.help_window.destroy()
        self.help_window = None

    def select_excel_file(self) -> None:
        file_path = filedialog.askopenfilename(title="選擇 Excel 檔案", filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")])
        if not file_path:
            return
        self.file_path_var.set(file_path)
        self.output_path_var.set(default_output_path(file_path))
        self._append_log("[INFO] 正在載入工作表名稱...")
        threading.Thread(target=self.load_worksheet_names, args=(file_path,), daemon=True).start()

    def load_worksheet_names(self, file_path: str) -> None:
        try:
            names = ExcelProcessor().load_worksheet_names(file_path)
            self.message_queue.put(("sheets", names))
        except Exception as exc:
            self.message_queue.put(("error", str(exc)))

    def select_output_file(self) -> None:
        initial = self.output_path_var.get() or default_output_path(self.file_path_var.get()) if self.file_path_var.get() else "整理完成.xlsx"
        file_path = filedialog.asksaveasfilename(title="選擇輸出檔案", initialfile=Path(initial).name, initialdir=str(Path(initial).parent), defaultextension=Path(initial).suffix or ".xlsx", filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")])
        if file_path:
            self.output_path_var.set(file_path)

    def start_processing(self) -> None:
        if os.path.abspath(self.file_path_var.get() or "") == os.path.abspath(self.output_path_var.get() or ""):
            if not messagebox.askyesno("覆蓋確認", "輸出路徑與原始檔相同，確定要覆蓋嗎？"):
                return
            self._append_log("[WARNING] 輸出路徑與原始檔相同，將覆蓋原始檔案。")
        self.start_button.configure(state="disabled")
        self.progress.set(0)
        self._append_log("[INFO] 開始整理...")
        args = (self.file_path_var.get(), self.worksheet_var.get(), self.range_var.get(), self.output_path_var.get())
        threading.Thread(target=self._worker_process, args=args, daemon=True).start()

    def _worker_process(self, file_path: str, worksheet_name: str, range_text: str, output_path: str) -> None:
        processor = ExcelProcessor(progress_callback=lambda value, maximum: self.message_queue.put(("progress", (value, maximum))), log_callback=lambda message: self.message_queue.put(("log", message)))
        try:
            stats = processor.process_excel_data(file_path, worksheet_name, range_text, output_path)
            self.message_queue.put(("done", stats))
        except Exception as exc:
            self.message_queue.put(("error", str(exc)))

    def _poll_queue(self) -> None:
        while not self.message_queue.empty():
            kind, payload = self.message_queue.get()
            if kind == "sheets":
                self.sheet_combo.configure(values=payload)
                if payload:
                    self.worksheet_var.set(payload[0])
                self._append_log("[OK] 工作表名稱載入完成。")
            elif kind == "progress":
                value, maximum = payload
                self.progress.set(value / maximum if maximum else 0)
            elif kind == "log":
                self._append_log("[INFO] " + payload)
            elif kind == "done":
                self.start_button.configure(state="normal")
                self.progress.set(1)
                message = payload.to_message()
                self._append_log("[OK] 整理完成！\n" + message)
                messagebox.showinfo("完成", "整理完成！\n\n" + message)
            elif kind == "error":
                self.start_button.configure(state="normal")
                self.show_error_message(payload)
        self.after(100, self._poll_queue)

    def _append_log(self, message: str) -> None:
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def show_error_message(self, message: str) -> None:
        self._append_log("[ERROR] " + message)
        messagebox.showerror("錯誤", message)


def main() -> None:
    """程式進入點。"""
    app = ExcelOrganizerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
