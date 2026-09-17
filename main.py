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

from config_manager import default_rules_path, load_rules, save_rules
from rule_engine import (
    evaluate_custom_level,
    find_override_pairs,
    find_rule_conflicts,
    is_blank_value,
)
from rule_models import (
    LABEL_TO_OPERATOR,
    OPERATOR_LABELS,
    VALID_GENDERS,
    CustomRule,
)


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


DEFAULT_RANGE = "M1:R2000"
APP_VERSION = "2.0"
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

CUSTOM_HELP_SECTIONS = [
    ("功能說明", "依 Excel 儲存格實際內容，按照使用者設定的規則，產生第一級至第四級。"),
    ("Excel 格式", "第一列必須為表頭，且必須有「管理級數」。使用男／女專用規則時，第一列也必須有「性別」，內容使用「男」或「女」；性別欄可以位於分析範圍外。"),
    ("分析資料範圍", "例如 M1:R2000 代表只分析 M～R 的來源欄位，資料由第 2 列開始。所有啟用規則的欄位都必須位於範圍內。"),
    ("規則種類", "支援區間、>、>=、<、<=、完全相符。區間包含上下限，例如 1～10 等同 1 <= value <= 10。"),
    ("比較範例", ">10.6 → 第三級\n<25.8 → 第二級\n數字或看起來是數字的文字都會嘗試轉換。"),
    ("完全相符範例", "++ → 第二級\n代謝症候群 → 第四級\n比較前會移除前後空白，但不是模糊搜尋。"),
    ("男女設定", "共用：男女皆使用。\n男／女：只套用相同性別。專用規則先判斷；若沒有符合，再使用共用規則。"),
    ("規則順序與衝突", "同一欄位與同一性別群組由上至下判斷，第一個符合的規則為準。可用上移／下移調整順序。新增或編輯時會精確檢查重疊；完全相同規則不得重複，其他重疊可由使用者確認後保存，列表會顯示 ⚠。開始整理前也會再次確認。男女專用規則與共用規則可能同時命中時只屬覆寫提醒，專用規則仍優先。"),
    ("注意事項", "空白儲存格會跳過。數值規則遇到無法轉換的文字不會判斷成功，也不會中斷。規則可存成 JSON；損壞的 JSON 不會阻止程式啟動。"),
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


@dataclass
class CustomProcessStats(ProcessStats):
    applied_rules: int = 0
    numeric_conversion_failures: int = 0
    unmatched_cells: int = 0

    def to_message(self) -> str:
        return (
            f"實際處理的工作表名稱：{self.worksheet_name}\n"
            f"實際處理的範圍：{self.actual_range}\n"
            f"掃描的資料列數：{self.scanned_rows}\n"
            f"掃描的非空白儲存格數：{self.non_empty_cells}\n"
            f"套用規則數：{self.applied_rules}\n"
            f"第一級項目數：{self.level_counts['第一級']}\n"
            f"第二級項目數：{self.level_counts['第二級']}\n"
            f"第三級項目數：{self.level_counts['第三級']}\n"
            f"第四級項目數：{self.level_counts['第四級']}\n"
            f"無法轉換成數字的項目數：{self.numeric_conversion_failures}\n"
            f"未符合任何規則的項目數：{self.unmatched_cells}\n"
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

    def load_worksheet_info(self, file_path: str, worksheet_name: str) -> dict[int, str]:
        """讀取指定工作表第一列表頭，供規則欄位下拉選單使用。"""
        self._validate_input_file(file_path)
        self._open_excel(file_path, read_only=True)
        try:
            worksheet = self._get_worksheet(worksheet_name)
            last_col = worksheet.Cells(1, worksheet.Columns.Count).End(-4159).Column
            return {
                col: str(worksheet.Cells(1, col).Value or "").strip()
                for col in range(1, last_col + 1)
            }
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

    def process_custom_data(self, file_path: str, worksheet_name: str, range_text: str, output_path: str, rules: list[CustomRule]) -> CustomProcessStats:
        """依儲存格內容與自訂規則整理，重用 V1 的欄位建立與輸出流程。"""
        original_range = validate_range(range_text)
        self._validate_paths(file_path, output_path)
        enabled_rules = [rule for rule in rules if rule.enabled]
        if not enabled_rules:
            raise ExcelProcessError("請至少建立並啟用一條自訂規則。")
        for rule in enabled_rules:
            try:
                rule.validate()
            except ValueError as exc:
                raise ExcelProcessError(f"規則欄位 {rule.column}：{exc}") from exc
            col = column_letter_to_index(rule.column)
            if not original_range.start_col <= col <= original_range.end_col:
                raise ExcelProcessError(f"規則欄位 {rule.column} 不在目前分析資料範圍內。")

        self._open_excel(file_path, read_only=False)
        try:
            worksheet = self._get_worksheet(worksheet_name)
            source_headers = self._capture_source_headers(worksheet, original_range)
            management_col = self.find_management_level_column(worksheet)
            gender_col = self._find_header_column(worksheet, "性別")
            if any(rule.gender != "共用" for rule in enabled_rules) and gender_col is None:
                raise ExcelProcessError("找不到表頭「性別」，目前規則包含男女分別設定，請確認 Excel 第一列。")
            self._ensure_writable(output_path)
            level_columns, inserted_count = self.create_level_columns(worksheet, management_col)
            adjusted_range, adjusted_headers = self._adjust_source_range(original_range, source_headers, management_col, inserted_count)
            self._validate_adjusted_headers(adjusted_headers)
            if gender_col is not None and inserted_count and gender_col >= management_col + 1:
                gender_col += inserted_count
            rule_map: dict[int, list[CustomRule]] = {}
            for rule in enabled_rules:
                original_col = column_letter_to_index(rule.column)
                adjusted_col = original_col + inserted_count if inserted_count and original_col >= management_col + 1 else original_col
                rule_map.setdefault(adjusted_col, []).append(rule)
            self._clear_old_results(worksheet, level_columns, adjusted_range.end_row)
            self.log(f"載入 {len(enabled_rules)} 條規則")
            self.log(f"分析範圍：{range_text}")
            if gender_col is not None:
                self.log(f"找到性別欄：{column_index_to_letter(gender_col)}")
            stats = self._scan_custom_and_write(worksheet, worksheet_name, adjusted_range, adjusted_headers, level_columns, gender_col, rule_map)
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

    def _find_header_column(self, worksheet, header_name: str) -> int | None:
        last_col = worksheet.Cells(1, worksheet.Columns.Count).End(-4159).Column
        for col in range(1, last_col + 1):
            if str(worksheet.Cells(1, col).Value or "").strip() == header_name:
                return col
        return None

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

    def _scan_custom_and_write(self, worksheet, worksheet_name: str, excel_range: ExcelRange, headers: list[tuple[int, str]], level_columns: dict[str, int], gender_col: int | None, rule_map: dict[int, list[CustomRule]]) -> CustomProcessStats:
        stats = CustomProcessStats(worksheet_name=worksheet_name, actual_range=format_range(excel_range), scanned_rows=excel_range.data_row_count, applied_rules=sum(len(items) for items in rule_map.values()))
        maximum = max(1, excel_range.data_row_count)
        for row_index, row in enumerate(range(2, excel_range.end_row + 1), start=1):
            row_results = {header: [] for header in LEVEL_HEADERS}
            seen = {header: set() for header in LEVEL_HEADERS}
            gender = worksheet.Cells(row, gender_col).Value if gender_col else None
            for col, source_header in headers:
                if col not in rule_map:
                    continue
                value = worksheet.Cells(row, col).Value
                if is_blank_value(value):
                    continue
                stats.non_empty_cells += 1
                result = evaluate_custom_level(value, gender, rule_map[col])
                if result.numeric_conversion_failed:
                    stats.numeric_conversion_failures += 1
                if result.level is None:
                    stats.unmatched_cells += 1
                    continue
                if source_header not in seen[result.level]:
                    seen[result.level].add(source_header)
                    row_results[result.level].append(source_header)
                    stats.level_counts[result.level] += 1
            for level_header, names in row_results.items():
                if names:
                    worksheet.Cells(row, level_columns[level_header]).Value = "".join(f"{name}；" for name in names)
            if row_index % 10 == 0 or row_index == maximum:
                self.update_progress(row_index, maximum)
        return stats


class ExcelOrganizerApp(ctk.CTk):
    """V2.0 雙模式 CustomTkinter GUI 主視窗。"""

    def __init__(self) -> None:
        super().__init__()
        self.title(f"Excel 管理級數整理工具 V{APP_VERSION}")
        self.geometry("1100x820")
        self.minsize(900, 700)
        self.message_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.file_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.range_var = tk.StringVar(value=DEFAULT_RANGE)
        self.worksheet_var = tk.StringVar()
        self.rules_path = default_rules_path()
        self.rules: list[CustomRule] = []
        self.worksheet_headers: dict[int, str] = {}
        self.help_window: ctk.CTkToplevel | None = None
        self._build_ui()
        self._load_default_rules()
        self.after(100, self._poll_queue)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(16, 8))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Excel 管理級數整理工具", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(header, text=f"V{APP_VERSION}", font=ctk.CTkFont(size=16, weight="bold"), text_color=("#1f6aa5", "#3b8ed0")).grid(row=0, column=1, padx=12)
        ctk.CTkButton(header, text="使用說明", width=100, command=self.show_help_window).grid(row=0, column=2)

        common = ctk.CTkFrame(self, corner_radius=12)
        common.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 10))
        common.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(common, text="Excel 檔案", anchor="w").grid(row=0, column=0, padx=(18, 10), pady=(14, 6))
        ctk.CTkEntry(common, textvariable=self.file_path_var).grid(row=0, column=1, sticky="ew", pady=(14, 6))
        ctk.CTkButton(common, text="選擇檔案", width=110, command=self.select_excel_file).grid(row=0, column=2, padx=(10, 18), pady=(14, 6))
        ctk.CTkLabel(common, text="工作表", anchor="w").grid(row=1, column=0, padx=(18, 10), pady=6)
        self.sheet_combo = ctk.CTkComboBox(common, variable=self.worksheet_var, values=[], state="readonly", command=self._worksheet_changed)
        self.sheet_combo.grid(row=1, column=1, sticky="ew", pady=6)
        ctk.CTkLabel(common, text="輸出檔案", anchor="w").grid(row=2, column=0, padx=(18, 10), pady=(6, 14))
        ctk.CTkEntry(common, textvariable=self.output_path_var).grid(row=2, column=1, sticky="ew", pady=(6, 14))
        ctk.CTkButton(common, text="選擇輸出位置", width=110, command=self.select_output_file).grid(row=2, column=2, padx=(10, 18), pady=(6, 14))

        self.tabs = ctk.CTkTabview(self, command=self._tab_changed)
        self.tabs.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 10))
        color_tab = self.tabs.add("依顏色分級")
        custom_tab = self.tabs.add("自訂分級")
        self._build_color_tab(color_tab)
        self._build_custom_tab(custom_tab)

        activity = ctk.CTkFrame(self, corner_radius=12)
        activity.grid(row=3, column=0, sticky="nsew", padx=24, pady=(0, 18))
        activity.grid_columnconfigure(0, weight=1)
        activity.grid_rowconfigure(3, weight=1)
        self.start_button = ctk.CTkButton(activity, text="開始依顏色分級", height=42, font=ctk.CTkFont(size=15, weight="bold"), command=self.start_processing)
        self.start_button.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 8))
        self.progress = ctk.CTkProgressBar(activity, mode="determinate")
        self.progress.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 10)); self.progress.set(0)
        ctk.CTkLabel(activity, text="執行紀錄", font=ctk.CTkFont(size=14, weight="bold")).grid(row=2, column=0, sticky="w", padx=18)
        self.log_text = ctk.CTkTextbox(activity, wrap="word", height=130)
        self.log_text.grid(row=3, column=0, sticky="nsew", padx=18, pady=(6, 14))

    def _build_color_tab(self, tab) -> None:
        tab.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(tab, text="依 Excel 原生 ColorIndex 分級（保留 V1 判斷方式）").grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 5))
        ctk.CTkLabel(tab, text="分析資料範圍").grid(row=1, column=0, sticky="w", padx=12, pady=(3, 10))
        ctk.CTkEntry(tab, textvariable=self.range_var).grid(row=1, column=1, sticky="ew", pady=(3, 10))
        ctk.CTkLabel(tab, text="例：M1:R2000", text_color=("gray45", "gray65")).grid(row=1, column=2, padx=12, pady=(3, 10))

    def _build_custom_tab(self, tab) -> None:
        tab.grid_columnconfigure(0, weight=1)
        toolbar = ctk.CTkFrame(tab, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", padx=8, pady=(6, 3)); toolbar.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(toolbar, text="分析資料範圍").grid(row=0, column=0, padx=(4, 8))
        ctk.CTkEntry(toolbar, textvariable=self.range_var).grid(row=0, column=1, sticky="ew")
        for index, (text, command) in enumerate((("新增規則", self.add_rule), ("儲存規則", self.save_rules_dialog), ("載入規則", self.load_rules_dialog)), start=2):
            ctk.CTkButton(toolbar, text=text, width=90, command=command).grid(row=0, column=index, padx=(8, 0))
        ctk.CTkLabel(tab, text="分級規則設定（專用性別優先；同群組由上至下 first match wins）", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, sticky="w", padx=12, pady=(2, 3))
        self.rules_frame = ctk.CTkScrollableFrame(tab, height=190)
        self.rules_frame.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 6))
        self.rules_frame.grid_columnconfigure(3, weight=1)

    def _current_mode(self) -> str:
        return self.tabs.get()

    def _tab_changed(self) -> None:
        self.start_button.configure(text="開始自訂分級" if self._current_mode() == "自訂分級" else "開始依顏色分級")

    def _load_default_rules(self) -> None:
        try:
            self.rules = load_rules(self.rules_path)
        except ValueError as exc:
            self.rules = []
            self.after(200, lambda: messagebox.showwarning("規則載入警告", f"{exc}\n將以空白規則啟動。"))
        self._refresh_rules()

    def _refresh_rules(self) -> None:
        for child in self.rules_frame.winfo_children(): child.destroy()
        conflict_indexes = {
            index
            for conflict in find_rule_conflicts(self.rules)
            for index in (conflict.first_index, conflict.second_index)
        }
        headings = ("啟用", "欄位", "表頭", "性別", "判斷方式", "條件", "分級", "操作")
        for col, heading in enumerate(headings):
            ctk.CTkLabel(self.rules_frame, text=heading, font=ctk.CTkFont(weight="bold")).grid(row=0, column=col, padx=5, pady=4, sticky="w")
        if not self.rules:
            ctk.CTkLabel(self.rules_frame, text="尚未建立規則，請按「新增規則」。", text_color=("gray45", "gray65")).grid(row=1, column=0, columnspan=8, pady=18)
            return
        for row, rule in enumerate(self.rules, start=1):
            enabled = tk.BooleanVar(value=rule.enabled)
            ctk.CTkCheckBox(self.rules_frame, text="", width=24, variable=enabled, command=lambda r=rule, v=enabled: self._toggle_rule(r, v)).grid(row=row, column=0, padx=5)
            col_index = column_letter_to_index(rule.column)
            header = self.worksheet_headers.get(col_index, "") or "—"
            column_text = f"⚠ {rule.column}" if row - 1 in conflict_indexes else rule.column
            values = (column_text, header, rule.gender, OPERATOR_LABELS[rule.operator], rule.condition_text, rule.level)
            for col, value in enumerate(values, start=1): ctk.CTkLabel(self.rules_frame, text=str(value), anchor="w").grid(row=row, column=col, padx=5, pady=3, sticky="w")
            buttons = ctk.CTkFrame(self.rules_frame, fg_color="transparent")
            buttons.grid(row=row, column=7, padx=3)
            for text, command in (("編輯", lambda i=row-1: self.edit_rule(i)), ("刪除", lambda i=row-1: self.delete_rule(i)), ("↑", lambda i=row-1: self.move_rule(i, -1)), ("↓", lambda i=row-1: self.move_rule(i, 1))):
                ctk.CTkButton(buttons, text=text, width=42, height=25, command=command).pack(side="left", padx=2)

    def _toggle_rule(self, rule: CustomRule, variable: tk.BooleanVar) -> None:
        rule.enabled = variable.get(); self._auto_save(); self._refresh_rules()

    def add_rule(self) -> None: self._show_rule_editor(None)
    def edit_rule(self, index: int) -> None: self._show_rule_editor(index)

    def delete_rule(self, index: int) -> None:
        if messagebox.askyesno("刪除規則", "確定刪除此規則？"):
            self.rules.pop(index); self._auto_save(); self._refresh_rules()

    def move_rule(self, index: int, delta: int) -> None:
        target = index + delta
        if 0 <= target < len(self.rules):
            self.rules[index], self.rules[target] = self.rules[target], self.rules[index]
            self._auto_save(); self._refresh_rules()

    def _column_options(self) -> list[str]:
        if self.worksheet_headers:
            return [f"{column_index_to_letter(col)} - {header or '(空白表頭)'}" for col, header in self.worksheet_headers.items()]
        return [column_index_to_letter(col) for col in range(1, 53)]

    def _show_rule_editor(self, index: int | None) -> None:
        current = self.rules[index] if index is not None else None
        window = ctk.CTkToplevel(self); window.title("編輯規則" if current else "新增規則"); window.geometry("480x500"); window.transient(self); window.grab_set()
        window.grid_columnconfigure(1, weight=1)
        column = tk.StringVar(value=(current.column if current else self._column_options()[0]))
        gender = tk.StringVar(value=current.gender if current else "共用")
        operator = tk.StringVar(value=OPERATOR_LABELS[current.operator] if current else "區間")
        level = tk.StringVar(value=current.level if current else "第一級")
        value = tk.StringVar(value="" if not current or current.value is None else str(current.value))
        minimum = tk.StringVar(value="" if not current or current.minimum is None else str(current.minimum))
        maximum = tk.StringVar(value="" if not current or current.maximum is None else str(current.maximum))
        ctk.CTkLabel(window, text="來源欄位").grid(row=0, column=0, padx=18, pady=(20, 8), sticky="w")
        ctk.CTkComboBox(window, variable=column, values=self._column_options()).grid(row=0, column=1, padx=18, pady=(20, 8), sticky="ew")
        ctk.CTkLabel(window, text="性別").grid(row=1, column=0, padx=18, pady=8, sticky="w"); ctk.CTkComboBox(window, variable=gender, values=list(VALID_GENDERS), state="readonly").grid(row=1, column=1, padx=18, pady=8, sticky="ew")
        ctk.CTkLabel(window, text="判斷方式").grid(row=2, column=0, padx=18, pady=8, sticky="w")
        condition = ctk.CTkFrame(window, fg_color="transparent"); condition.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12); condition.grid_columnconfigure(1, weight=1)
        def render(*_):
            for child in condition.winfo_children(): child.destroy()
            op = LABEL_TO_OPERATOR[operator.get()]
            if op == "range":
                ctk.CTkLabel(condition, text="下限").grid(row=0, column=0, padx=6, pady=6); ctk.CTkEntry(condition, textvariable=minimum).grid(row=0, column=1, sticky="ew", padx=6, pady=6)
                ctk.CTkLabel(condition, text="上限").grid(row=1, column=0, padx=6, pady=6); ctk.CTkEntry(condition, textvariable=maximum).grid(row=1, column=1, sticky="ew", padx=6, pady=6)
            else:
                ctk.CTkLabel(condition, text="內容" if op == "exact" else "數值").grid(row=0, column=0, padx=6, pady=6); ctk.CTkEntry(condition, textvariable=value).grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        ctk.CTkComboBox(window, variable=operator, values=list(LABEL_TO_OPERATOR), state="readonly", command=render).grid(row=2, column=1, padx=18, pady=8, sticky="ew"); render()
        ctk.CTkLabel(window, text="分級").grid(row=4, column=0, padx=18, pady=8, sticky="w"); ctk.CTkComboBox(window, variable=level, values=LEVEL_HEADERS, state="readonly").grid(row=4, column=1, padx=18, pady=8, sticky="ew")
        def save() -> None:
            try:
                column_name = column.get().split(" - ", 1)[0].strip().upper()
                item = CustomRule(column_name, gender.get(), LABEL_TO_OPERATOR[operator.get()], level.get(), enabled=current.enabled if current else True, value=value.get(), minimum=minimum.get(), maximum=maximum.get()); item.validate()
            except (ValueError, KeyError) as exc:
                messagebox.showerror("規則錯誤", str(exc), parent=window); return
            candidate_rules = list(self.rules)
            if index is None: candidate_rules.append(item); saved_index = len(candidate_rules) - 1
            else: candidate_rules[index] = item; saved_index = index
            related = [conflict for conflict in find_rule_conflicts(candidate_rules) if saved_index in (conflict.first_index, conflict.second_index)]
            duplicate = next((conflict for conflict in related if conflict.is_duplicate), None)
            if duplicate:
                messagebox.showerror("此規則已存在", "相同欄位、性別、判斷方式、條件與分級的規則已存在，請修改目前規則。", parent=window)
                return
            override_count = sum(saved_index in pair for pair in find_override_pairs(candidate_rules))
            if related or override_count:
                details = self._format_conflict_message(candidate_rules, saved_index, related, override_count)
                if not self._ask_confirmation("偵測到規則可能重疊", details, "仍然儲存", "返回修改", window):
                    return
            self.rules = candidate_rules
            self._auto_save(); self._refresh_rules(); window.destroy()
        actions = ctk.CTkFrame(window, fg_color="transparent"); actions.grid(row=5, column=0, columnspan=2, pady=20)
        ctk.CTkButton(actions, text="儲存", command=save).pack(side="left", padx=8); ctk.CTkButton(actions, text="取消", fg_color="gray50", command=window.destroy).pack(side="left", padx=8)

    def _format_conflict_message(self, rules: list[CustomRule], saved_index: int, conflicts, override_count: int) -> str:
        current = rules[saved_index]
        header = self.worksheet_headers.get(column_letter_to_index(current.column), "") or current.column
        lines = [f"欄位：{header}", f"目前規則：{current.condition_text} → {current.level}"]
        for conflict in conflicts:
            other_index = conflict.second_index if conflict.first_index == saved_index else conflict.first_index
            other = rules[other_index]
            lines.append(f"衝突規則：{other.condition_text} → {other.level}")
            if conflict.is_same_condition:
                lines.append("原因：相同條件被設定為不同分級。")
            else:
                lines.append("原因：兩條規則存在可同時符合的值。")
        if override_count:
            lines.append(f"覆寫提醒：另有 {override_count} 條同欄位的男女專用／共用規則可能同時命中；專用規則會優先。")
        lines.append("實際分析時將依規則順序，使用第一個符合的規則。")
        return "\n".join(lines)

    def _ask_confirmation(self, title: str, message: str, confirm_text: str, cancel_text: str, parent=None) -> bool:
        """顯示具有明確動作文字的 modal 確認視窗。"""
        dialog = ctk.CTkToplevel(parent or self)
        dialog.title(title)
        dialog.geometry("620x430")
        dialog.minsize(500, 320)
        dialog.transient(parent or self)
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(0, weight=1)
        result = {"confirmed": False}
        text = ctk.CTkTextbox(dialog, wrap="word", corner_radius=8)
        text.grid(row=0, column=0, sticky="nsew", padx=18, pady=(18, 10))
        text.insert("end", message)
        text.configure(state="disabled")
        actions = ctk.CTkFrame(dialog, fg_color="transparent")
        actions.grid(row=1, column=0, pady=(4, 18))

        def close(confirmed: bool = False) -> None:
            result["confirmed"] = confirmed
            dialog.grab_release()
            dialog.destroy()

        ctk.CTkButton(actions, text=confirm_text, command=lambda: close(True)).pack(side="left", padx=8)
        ctk.CTkButton(actions, text=cancel_text, fg_color="gray50", command=close).pack(side="left", padx=8)
        dialog.protocol("WM_DELETE_WINDOW", close)
        dialog.grab_set()
        dialog.wait_window()
        if parent is not None and parent.winfo_exists():
            parent.grab_set()
        return result["confirmed"]

    def _auto_save(self) -> None:
        try: save_rules(self.rules_path, self.rules)
        except OSError as exc: self._append_log(f"[WARNING] 自動儲存規則失敗：{exc}")

    def save_rules_dialog(self) -> None:
        path = filedialog.asksaveasfilename(title="儲存規則", defaultextension=".json", filetypes=[("JSON", "*.json")], initialfile="custom_rules.json")
        if path:
            try: save_rules(path, self.rules); self._append_log(f"[OK] 規則已儲存：{path}")
            except OSError as exc: self.show_error_message(str(exc))

    def load_rules_dialog(self) -> None:
        path = filedialog.askopenfilename(title="載入規則", filetypes=[("JSON", "*.json")])
        if path:
            try: self.rules = load_rules(path); self._auto_save(); self._refresh_rules(); self._append_log(f"[OK] 已載入 {len(self.rules)} 條規則。")
            except ValueError as exc: self.show_error_message(str(exc))

    def show_help_window(self) -> None:
        if self.help_window is not None and self.help_window.winfo_exists(): self.help_window.lift(); return
        sections = CUSTOM_HELP_SECTIONS if self._current_mode() == "自訂分級" else HELP_SECTIONS
        self.help_window = ctk.CTkToplevel(self); self.help_window.title(f"{self._current_mode()}－使用說明"); self.help_window.geometry("720x620"); self.help_window.transient(self)
        self.help_window.protocol("WM_DELETE_WINDOW", self.close_help_window); self.help_window.columnconfigure(0, weight=1); self.help_window.rowconfigure(0, weight=1)
        text = ctk.CTkTextbox(self.help_window, wrap="word"); text.grid(row=0, column=0, sticky="nsew", padx=16, pady=(16, 8))
        for heading, body in sections: text.insert("end", f"【{heading}】\n{body}\n\n")
        text.configure(state="disabled"); ctk.CTkButton(self.help_window, text="關閉", command=self.close_help_window).grid(row=1, column=0, pady=(4, 14)); self.help_window.grab_set()

    def close_help_window(self) -> None:
        if self.help_window is not None:
            try: self.help_window.grab_release()
            except (tk.TclError, RuntimeError): pass
            self.help_window.destroy(); self.help_window = None

    def select_excel_file(self) -> None:
        path = filedialog.askopenfilename(title="選擇 Excel 檔案", filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")])
        if not path: return
        self.file_path_var.set(path); self.output_path_var.set(default_output_path(path)); self._append_log("[INFO] 正在載入工作表名稱...")
        threading.Thread(target=self.load_worksheet_names, args=(path,), daemon=True).start()

    def load_worksheet_names(self, path: str) -> None:
        try: self.message_queue.put(("sheets", ExcelProcessor().load_worksheet_names(path)))
        except Exception as exc: self.message_queue.put(("error", str(exc)))

    def _worksheet_changed(self, _name: str | None = None) -> None:
        if self.file_path_var.get() and self.worksheet_var.get(): threading.Thread(target=self._load_headers_worker, daemon=True).start()

    def _load_headers_worker(self) -> None:
        try: self.message_queue.put(("headers", ExcelProcessor().load_worksheet_info(self.file_path_var.get(), self.worksheet_var.get())))
        except Exception as exc: self.message_queue.put(("error", str(exc)))

    def select_output_file(self) -> None:
        initial = self.output_path_var.get() or (default_output_path(self.file_path_var.get()) if self.file_path_var.get() else "整理完成.xlsx")
        path = filedialog.asksaveasfilename(title="選擇輸出檔案", initialfile=Path(initial).name, initialdir=str(Path(initial).parent), defaultextension=Path(initial).suffix or ".xlsx", filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")])
        if path: self.output_path_var.set(path)

    def start_processing(self) -> None:
        mode = self._current_mode()
        if mode == "自訂分級":
            conflicts = find_rule_conflicts(self.rules)
            if conflicts and not self._ask_confirmation(
                "規則重疊確認",
                f"目前共有 {len(conflicts)} 組可能重疊的規則。\n\n"
                "繼續分析時將依規則排列順序，使用第一個符合的規則。\n\n"
                "請確認是否繼續分析。",
                "繼續分析",
                "取消",
            ):
                return
        if os.path.abspath(self.file_path_var.get() or "") == os.path.abspath(self.output_path_var.get() or ""):
            if not messagebox.askyesno("覆蓋確認", "輸出路徑與原始檔相同，確定要覆蓋嗎？"): return
        self.start_button.configure(state="disabled"); self.progress.set(0)
        self._append_log(f"[INFO] 開始{mode}...")
        args = (mode, self.file_path_var.get(), self.worksheet_var.get(), self.range_var.get(), self.output_path_var.get(), list(self.rules))
        threading.Thread(target=self._worker_process, args=args, daemon=True).start()

    def _worker_process(self, mode: str, file_path: str, worksheet_name: str, range_text: str, output_path: str, rules: list[CustomRule]) -> None:
        processor = ExcelProcessor(lambda value, maximum: self.message_queue.put(("progress", (value, maximum))), lambda message: self.message_queue.put(("log", message)))
        try:
            stats = processor.process_custom_data(file_path, worksheet_name, range_text, output_path, rules) if mode == "自訂分級" else processor.process_excel_data(file_path, worksheet_name, range_text, output_path)
            self.message_queue.put(("done", (mode, stats)))
        except Exception as exc: self.message_queue.put(("error", str(exc)))

    def _poll_queue(self) -> None:
        while not self.message_queue.empty():
            kind, payload = self.message_queue.get()
            if kind == "sheets":
                self.sheet_combo.configure(values=payload)
                if payload: self.worksheet_var.set(payload[0]); self._worksheet_changed(payload[0])
                self._append_log("[OK] 工作表名稱載入完成。")
            elif kind == "headers": self.worksheet_headers = payload; self._refresh_rules(); self._append_log("[OK] 第一列表頭載入完成。")
            elif kind == "progress":
                value, maximum = payload; self.progress.set(value / maximum if maximum else 0)
            elif kind == "log": self._append_log("[INFO] " + payload)
            elif kind == "done":
                mode, stats = payload; self.start_button.configure(state="normal"); self.progress.set(1); message = stats.to_message(); self._append_log(f"[OK] {mode}完成！\n{message}"); messagebox.showinfo("完成", f"{mode}完成！\n\n{message}")
            elif kind == "error": self.start_button.configure(state="normal"); self.show_error_message(payload)
        self.after(100, self._poll_queue)

    def _append_log(self, message: str) -> None:
        self.log_text.insert("end", message + "\n"); self.log_text.see("end")

    def show_error_message(self, message: str) -> None:
        self._append_log("[ERROR] " + message); messagebox.showerror("錯誤", message)

def main() -> None:
    """程式進入點。"""
    app = ExcelOrganizerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
