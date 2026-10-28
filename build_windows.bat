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
pyinstaller --noconfirm --clean --onefile --windowed --name ExcelSheetMerger --paths src --collect-submodules openpyxl --collect-submodules xlrd --collect-all ttkbootstrap app.py
if errorlevel 1 goto FAIL
if not exist dist\ExcelSheetMerger.exe goto FAIL
python -m pip freeze > dist\BUILD_DEPENDENCIES.txt
powershell -NoProfile -Command "$r = Join-Path $env:TEMP 'excel-merger-local-smoke.txt'; $p = Start-Process -FilePath (Resolve-Path 'dist\ExcelSheetMerger.exe') -ArgumentList @('--self-test', $r) -Wait -PassThru; if ($p.ExitCode -ne 0 -or -not (Test-Path $r)) { exit 1 }; if ((Get-Content $r -Raw) -notmatch '^PASS:') { exit 1 }"
if errorlevel 1 goto FAIL
echo.
echo [OK] dist\ExcelSheetMerger.exe
pause
exit /b 0
:FAIL
echo [ERROR] Build or test failed. See logs above.
pause
exit /b 1
