@echo off
setlocal
chcp 65001 >nul

cd /d "%~dp0"
set "PROJECT_DIR=%CD%"
set "APP_NAME=Excel管理級數整理工具"
set "SPEC_FILE=%PROJECT_DIR%\Excel管理級數整理工具.spec"
set "DIST_EXE=%PROJECT_DIR%\dist\%APP_NAME%.exe"
set "RELEASE_DIR=%PROJECT_DIR%\release"
set "RELEASE_EXE=%RELEASE_DIR%\%APP_NAME%.exe"

echo ========================================
echo  Excel 管理級數整理工具 - EXE 自動打包
echo ========================================
echo [INFO] 目前工作目錄：
echo "%PROJECT_DIR%"
echo.

if not exist "%PROJECT_DIR%\main.py" (
    echo [ERROR] 找不到 main.py
    goto :failed
)
echo [OK] 已找到 main.py

set "SYSTEM_PYTHON="
python --version >nul 2>&1
if not errorlevel 1 set "SYSTEM_PYTHON=python"

if not defined SYSTEM_PYTHON (
    py --version >nul 2>&1
    if not errorlevel 1 set "SYSTEM_PYTHON=py"
)

if not defined SYSTEM_PYTHON (
    echo [ERROR] 找不到 Python，請確認 Python 已安裝並加入 PATH。
    goto :failed
)

if exist "%PROJECT_DIR%\.venv\Scripts\python.exe" (
    set "PYTHON=%PROJECT_DIR%\.venv\Scripts\python.exe"
    echo [INFO] 偵測到專案虛擬環境，將優先使用 .venv。
) else (
    set "PYTHON=%SYSTEM_PYTHON%"
    echo [INFO] 未偵測到 .venv，將使用系統 Python。
)

echo [INFO] Python:
"%PYTHON%" -c "import sys; print(sys.executable)"
if errorlevel 1 (
    echo [ERROR] 無法取得 Python 路徑。
    goto :failed
)
"%PYTHON%" --version
if errorlevel 1 (
    echo [ERROR] 無法取得 Python 版本。
    goto :failed
)
echo [OK] Python 可以使用。
echo.

if exist "%PROJECT_DIR%\requirements.txt" (
    echo [INFO] 正在安裝 requirements.txt 中的相依套件...
    "%PYTHON%" -m pip install -r "%PROJECT_DIR%\requirements.txt"
    if errorlevel 1 (
        echo [ERROR] requirements.txt 套件安裝失敗
        goto :failed
    )
    echo [OK] requirements.txt 套件安裝完成。
) else (
    echo [WARNING] 找不到 requirements.txt，將略過相依套件安裝。
)
echo.

echo [INFO] 正在檢查 PyInstaller...
"%PYTHON%" -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo [WARNING] 找不到 PyInstaller，正在自動安裝...
    "%PYTHON%" -m pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] PyInstaller 安裝失敗
        goto :failed
    )
    "%PYTHON%" -m PyInstaller --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] PyInstaller 安裝失敗
        goto :failed
    )
)
echo [OK] PyInstaller 可以使用：
"%PYTHON%" -m PyInstaller --version
if errorlevel 1 (
    echo [ERROR] PyInstaller 版本檢查失敗
    goto :failed
)
echo.

if not exist "%SPEC_FILE%" (
    echo [ERROR] 找不到 Excel管理級數整理工具.spec
    goto :failed
)
echo [OK] 已找到 PyInstaller spec 檔案。

echo [INFO] 正在清理舊的 build 與 dist 資料夾...
if exist "%PROJECT_DIR%\build" rmdir /s /q "%PROJECT_DIR%\build"
if exist "%PROJECT_DIR%\build" (
    echo [ERROR] 無法清除 build 資料夾，請確認其中檔案未被占用。
    goto :failed
)
if exist "%PROJECT_DIR%\dist" rmdir /s /q "%PROJECT_DIR%\dist"
if exist "%PROJECT_DIR%\dist" (
    echo [ERROR] 無法清除 dist 資料夾，請確認其中檔案未被占用。
    goto :failed
)
if exist "%RELEASE_EXE%" del /f /q "%RELEASE_EXE%"
if exist "%RELEASE_EXE%" (
    echo [ERROR] 無法刪除舊的 release\%APP_NAME%.exe，請確認檔案未被占用。
    goto :failed
)
echo [OK] 舊打包檔案清理完成，release 中的其他檔案不受影響。
echo.

echo [BUILD] 開始執行 PyInstaller...
"%PYTHON%" -m PyInstaller --noconfirm --clean "%SPEC_FILE%"
if errorlevel 1 (
    echo.
    echo ========================================
    echo [ERROR] EXE 打包失敗
    echo ========================================
    echo 請查看上方錯誤訊息。
    goto :failed
)
echo [OK] PyInstaller 已成功執行。

if not exist "%DIST_EXE%" (
    echo [ERROR] PyInstaller 執行完成，但找不到 EXE。
    echo [INFO] 預期路徑：
    echo "%DIST_EXE%"
    goto :failed
)
echo [OK] 已找到 dist 輸出的 EXE。

if not exist "%RELEASE_DIR%" (
    mkdir "%RELEASE_DIR%"
    if errorlevel 1 (
        echo [ERROR] 無法建立 release 資料夾。
        goto :failed
    )
)

copy /y "%DIST_EXE%" "%RELEASE_EXE%" >nul
if errorlevel 1 (
    echo [ERROR] 無法將 EXE 複製到 release 資料夾。
    goto :failed
)

if not exist "%RELEASE_EXE%" (
    echo [ERROR] release 未成功取得 EXE。
    goto :failed
)

for %%I in ("%RELEASE_EXE%") do set "EXE_SIZE=%%~zI"
if "%EXE_SIZE%"=="" (
    echo [ERROR] 無法取得 EXE 檔案大小。
    goto :failed
)
if %EXE_SIZE% LEQ 0 (
    echo [ERROR] EXE 檔案大小異常：0 bytes
    goto :failed
)

echo [OK] release 已成功取得 EXE，且檔案大小大於 0。
echo [INFO] 正式 EXE 已驗證，正在清除本次打包產生的 build 與 dist 資料夾...
if exist "%PROJECT_DIR%\build" rmdir /s /q "%PROJECT_DIR%\build"
if exist "%PROJECT_DIR%\build" (
    echo [ERROR] EXE 已建立，但無法清除 build 資料夾，請確認其中檔案未被占用。
    goto :failed
)
if exist "%PROJECT_DIR%\dist" rmdir /s /q "%PROJECT_DIR%\dist"
if exist "%PROJECT_DIR%\dist" (
    echo [ERROR] EXE 已建立，但無法清除 dist 資料夾，請確認其中檔案未被占用。
    goto :failed
)
echo [OK] build 與 dist 資料夾已清除，只保留 release 中的正式 EXE。
echo.
echo ========================================
echo [SUCCESS] EXE 打包完成
echo ========================================
echo [INFO] 完整路徑：
echo "%RELEASE_EXE%"
echo [INFO] EXE 檔案大小：%EXE_SIZE% bytes
goto :success

:failed
echo.
echo [ERROR] 打包流程已停止，請依上方階段訊息排除問題。
echo [INFO] 若 PyInstaller 已開始執行，build 與 dist 將保留供除錯。
echo.
echo 按任意鍵關閉視窗...
pause >nul
endlocal
exit /b 1

:success
echo.
echo 按任意鍵關閉視窗...
pause >nul
endlocal
exit /b 0
