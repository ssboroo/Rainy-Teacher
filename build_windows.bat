@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo Багшийн Туслах Windows хувилбар бэлдэж байна...
py -m pip install --upgrade pyinstaller
if errorlevel 1 goto failed
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "БагшийнТуслах" bagshiin_tuslah.py
if errorlevel 1 goto failed
echo.
echo Амжилттай. Файл: dist\БагшийнТуслах.exe
pause
exit /b 0
:failed
echo Алдаа гарлаа. Python болон pip суусныг шалгана уу.
pause
exit /b 1
