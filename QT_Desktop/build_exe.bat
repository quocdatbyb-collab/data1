@echo off
chcp 65001 >nul
REM ==========================================================
REM  Đóng gói QT Quyết Toán thành 1 file .exe (chạy trên Windows)
REM  Yêu cầu: đã cài Python 3.8+ (tick "Add Python to PATH")
REM  Tác giả: 8702 - Nguyễn Quốc Đạt
REM ==========================================================
cd /d "%~dp0"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller==6.10.0
python tests\run_all.py "%TEMP%\qt_selftest" || goto :error
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name QT_QuyetToan --icon assets\app.ico --add-data "assets;assets" ^
  --version-file version_info.txt --exclude-module numpy --exclude-module pandas ^
  main.py || goto :error
echo.
echo Xong! File chuong trinh: dist\QT_QuyetToan.exe
pause
exit /b 0
:error
echo Co loi khi dong goi.
pause
exit /b 1
