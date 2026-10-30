@echo off
chcp 65001 >nul
title Cai dat QT Quyet Toan ra Desktop - Tac gia: 8702 - Nguyen Quoc Dat
cd /d "%~dp0"
echo ==========================================================
echo   CAI DAT "QT QUYET TOAN" RA MAN HINH DESKTOP
echo   Tac gia: 8702 - Nguyen Quoc Dat
echo ==========================================================
where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo May chua co Python. Hay cai Python 3.8 tro len tai https://www.python.org/downloads/
  echo Khi cai NHO TICK o "Add Python to PATH", sau do chay lai file nay.
  start https://www.python.org/downloads/
  pause
  exit /b 1
)
echo [1/3] Dang cai thu vien...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt pyinstaller==6.10.0 || goto :error
echo [2/3] Dang dong goi chuong trinh (mat 1-3 phut)...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name QT_QuyetToan --icon assets\app.ico --add-data "assets;assets" ^
  --version-file version_info.txt --exclude-module numpy --exclude-module pandas ^
  main.py >nul 2>build_log.txt || goto :error
echo [3/3] Dang dua ra Desktop...
for /f "usebackq delims=" %%D in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) do set "DESK=%%D"
if not exist "%DESK%\QT_QuyetToan" mkdir "%DESK%\QT_QuyetToan"
copy /y dist\QT_QuyetToan.exe "%DESK%\QT_QuyetToan\QT_QuyetToan.exe" >nul || goto :error
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%DESK%\QT Quyet Toan.lnk'); $s.TargetPath='%DESK%\QT_QuyetToan\QT_QuyetToan.exe'; $s.WorkingDirectory='%DESK%\QT_QuyetToan'; $s.Description='QT Quyet Toan - 8702 Nguyen Quoc Dat'; $s.Save()"
echo.
echo XONG! Tren Desktop da co bieu tuong "QT Quyet Toan".
start "" "%DESK%\QT_QuyetToan\QT_QuyetToan.exe"
pause
exit /b 0
:error
echo.
echo Co loi khi cai dat. Xem file build_log.txt va gui cho nguoi ho tro.
pause
exit /b 1
