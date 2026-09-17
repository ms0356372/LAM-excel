# Excel 管理級數整理工具 V2.0

Windows 桌面工具，使用 CustomTkinter 與 Excel COM（`pywin32`）整理健康檢查或管理級數資料。V2.0 延續 V1 的 ColorIndex 分級流程，並在同一個 GUI 新增可保存規則的「自訂分級」。程式需在已安裝 Microsoft Excel 的 Windows 電腦執行。

## 功能模式

### 1. 依顏色分級（V1 相容）

此模式仍直接讀取 Excel 原生 `cell.Interior.ColorIndex`，映射與 V1 完全相同：

| ColorIndex | 結果 |
|---:|---|
| -4142 | 第一級 |
| 34 | 第二級 |
| 6 | 第三級 |
| 3 | 第四級 |

其他顏色及空白儲存格不處理。結果依來源欄由左至右，以 `表頭；` 寫入第一級至第四級欄。舊的掃描函式、ColorIndex 取得方式與映射常數保留不變。

### 2. 自訂分級

依儲存格實際內容套用 GUI 中的規則，支援：

- 區間（包含上下限）
- `>`、`>=`、`<`、`<=`
- 完全相符（移除前後空白後 exact match，不做模糊搜尋）
- 共用、男、女三種條件
- 規則新增、編輯、刪除、啟用／停用、上移、下移
- JSON 自動保存、另存與載入

同一欄位先依資料列的性別測試專用規則；專用規則沒有命中時才測試共用規則。各群組依畫面順序採 **first match wins**。若性別空白或不是男／女，只會使用共用規則。含男／女規則時，第一列必須有完全相同的 `性別` 表頭；它可以位於分析範圍之外。

數值、整數及數字文字（含前後空白）都會轉成數值比較。空白直接跳過，無法轉換的文字視為不符合且計入統計。規則重疊可以保存，介面會提醒實際結果由規則順序決定。

## Excel 格式與輸出

1. 支援 `.xlsx`、`.xlsm`、`.xls`；`.xlsm` 會維持巨集格式。
2. 第一列必須是表頭，並包含完全相同的 `管理級數`。
3. 分析範圍必須採 `M1:R2000` 形式並由第 1 列開始。
4. 程式會在 `管理級數` 後建立或重用 `第一級`、`第二級`、`第三級`、`第四級`，執行前清除舊結果。
5. 自訂模式的每個啟用規則欄位都必須位於分析範圍，否則執行會停止並顯示欄位名稱。
6. 同列同級的項目依來源欄由左至右排列，同一表頭只寫入一次。
7. 建議先關閉活頁簿並備份原始檔；程式預設另存為 `原檔名_整理完成`。

## 規則保存

開發環境預設寫入專案根目錄的 `custom_rules.json`；打包版寫入 `%APPDATA%\Excel管理級數整理工具\custom_rules.json`，避免單檔 EXE 的暫存目錄問題。格式版本為 1，例如：

```json
{
  "version": 1,
  "rules": [
    {"column": "M", "gender": "共用", "operator": "range", "minimum": 1.0, "maximum": 10.0, "level": "第一級", "enabled": true},
    {"column": "N", "gender": "男", "operator": ">=", "value": 90.0, "level": "第三級", "enabled": true},
    {"column": "O", "gender": "共用", "operator": "exact", "value": "++", "level": "第二級", "enabled": true}
  ]
}
```

檔案不存在時以空規則啟動；JSON 損壞或內容無效時顯示警告，主程式仍會以空規則啟動。

## 安裝與執行

請使用 Windows、Python 3.10+ 並安裝 Microsoft Excel：

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

主畫面上方選擇 Excel、工作表和輸出位置，再切換「依顏色分級」或「自訂分級」頁籤。兩個頁籤各有對應的使用說明；下方共用進度條與執行紀錄。

## EXE 打包

直接雙擊：

```text
build_exe.bat
```

批次檔會安裝 `requirements.txt`、使用 `Excel管理級數整理工具.spec` 執行 PyInstaller、驗證成品並輸出：

```text
release\Excel管理級數整理工具.exe
```

`build`、`dist` 只在失敗時保留以供除錯。目標電腦仍需安裝 Microsoft Excel。

## 測試

跨平台的規則引擎與 JSON 測試：

```bash
python -m pytest -q
python -m py_compile main.py rule_models.py rule_engine.py config_manager.py
```

測試涵蓋區間邊界、嚴格比較、exact match、無效數字、空白、男女優先、共用 fallback、first match wins、停用規則、驗證及 JSON 損壞。Linux CI 無法啟動 Windows Excel COM，因此仍需在 Windows 實機驗證：

1. 四種 V1 ColorIndex 的輸出與既有檔案結果完全一致。
2. 新增四個結果欄及已存在時重用、清除舊結果。
3. 管理級數位於來源範圍前方時，插欄後來源欄位仍正確位移。
4. 性別欄位在分析範圍外仍能套用男女規則。
5. `.xlsx`、`.xlsm`、`.xls` 另存，尤其 `.xlsm` 巨集保留。
6. `build_exe.bat` 可產生並啟動 release EXE，規則可在重啟後載入。
