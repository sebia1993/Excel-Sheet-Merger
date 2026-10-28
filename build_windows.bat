@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python 3.12 64-bit is required to build the EXE.
  echo Run this on GitHub Actions instead if Python is not installed.
  pause
  exit /b 1
)
py -3.12 -m venv .venv
if errorlevel 1 goto FAIL
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto FAIL
python -m pip install -r requirements-build.txt
if errorlevel 1 goto FAIL
python -m unittest discover -s tests -v
if errorlevel 1 goto FAIL
pyinstaller --noconfirm --clean --onefile --windowed --name ExcelSheetMerger --paths src --collect-submodules openpyxl --collect-all ttkbootstrap app.py
if errorlevel 1 goto FAIL
if not exist dist\ExcelSheetMerger.exe goto FAIL
echo.
echo [OK] dist\ExcelSheetMerger.exe
pause
exit /b 0
:FAIL
echo [ERROR] Build or test failed. See logs above.
pause
exit /b 1
