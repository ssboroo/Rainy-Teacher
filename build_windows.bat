@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo Багшийн Туслах Windows хувилбар бэлдэж байна...
py -m pip install --upgrade pyinstaller pillow python-pptx pywin32
if errorlevel 1 goto failed
py prepare_assets.py
if errorlevel 1 goto failed
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "RainyTeacher" --icon "assets\rainy_teacher.ico" --add-data "assets\rainy_teacher_mark.png;assets" --add-data "assets\rainy_teacher_mark.png.b64;assets" --add-data "assets\rainy_teacher.ico;assets" bagshiin_tuslah.py
if errorlevel 1 goto failed
where iscc >nul 2>nul
if errorlevel 1 (
  echo EXE бэлэн боллоо. Суулгагч үүсгэхийн тулд Inno Setup суулгаад дахин ажиллуулна уу.
  echo https://jrsoftware.org/isinfo.php
  pause
  exit /b 0
)
iscc /DAppVersion=1.7.1 installer.iss
if errorlevel 1 goto failed
echo.
echo Амжилттай. EXE: dist\RainyTeacher.exe
echo Суулгагч: dist-installer\RainyTeacher-Setup-1.7.1.exe
pause
exit /b 0
:failed
echo Алдаа гарлаа. Python болон pip суусныг шалгана уу.
pause
exit /b 1
