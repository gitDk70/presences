@echo off
setlocal
chcp 65001 >nul

set APP_NAME=Presences

echo >>> Installing Python dependencies...
pip install openpyxl "qrcode[pil]" keyring pyinstaller --quiet
pip uninstall pathlib -y 2>nul

echo >>> Building %APP_NAME%.exe...
pyinstaller ^
  --onefile ^
  --windowed ^
  --clean ^
  --noconfirm ^
  --name "%APP_NAME%" ^
  --add-data "user.png;." ^
  --add-data "logo.png;." ^
  --collect-submodules keyring ^
  --icon "user.ico" ^
  presences.py

if errorlevel 1 (
    echo.
    echo ERROR: build failed. See output above.
    pause
    exit /b 1
)

echo.
echo === Done! ===
echo   EXE: dist\%APP_NAME%.exe
echo.
echo Double-click dist\%APP_NAME%.exe to run.
pause
