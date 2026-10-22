@echo off
setlocal
cd /d "%~dp0"

python -m pip install --quiet pyinstaller
if errorlevel 1 goto fail

python -m PyInstaller --noconfirm --clean --windowed --noupx ^
  --name ShortcutIconFixer ^
  --icon assets\app.ico ^
  --version-file assets\version_info.txt ^
  --add-data "assets\app.ico;assets" ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import win32timezone ^
  --hidden-import pefile ^
  app.py
if errorlevel 1 goto fail

echo.
echo Build finished: dist\ShortcutIconFixer.exe
pause
exit /b 0

:fail
echo.
echo Build FAILED.
pause
exit /b 1
