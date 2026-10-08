@echo off
chcp 65001 >nul
setlocal
title ติดตั้ง Tong Heer INVB
set "SRC=%~dp0app"
set "APPDIR=%LOCALAPPDATA%\TongHeerINVB"
set "PYDIR=%LOCALAPPDATA%\Programs\Python\Python312"
set "PY=%PYDIR%\python.exe"
set "PYW=%PYDIR%\pythonw.exe"

echo ============================================
echo   ติดตั้งโปรแกรม Tong Heer INVB
echo ============================================
echo.
if not exist "%SRC%\tongheer_app.pyw" (
  echo [ผิดพลาด] ไม่พบโฟลเดอร์ app ข้างไฟล์ติดตั้ง - ต้องคัดลอกมาทั้งโฟลเดอร์ TongHeer_INVB_Setup
  goto :fail
)

echo [1/4] ตรวจ Python ...
if exist "%PY%" goto :have_python
where winget >nul 2>nul
if errorlevel 1 (
  echo [ผิดพลาด] เครื่องนี้ไม่มี winget - ติดตั้ง Python 3.12 จาก https://www.python.org/downloads/ ก่อน แล้วรันไฟล์นี้อีกครั้ง
  goto :fail
)
echo       ยังไม่มี Python กำลังดาวน์โหลดและติดตั้ง (ประมาณ 25 MB) ...
winget install --id Python.Python.3.12 -e --scope user --silent --accept-package-agreements --accept-source-agreements
if not exist "%PY%" (
  echo [ผิดพลาด] ติดตั้ง Python ไม่สำเร็จ
  goto :fail
)
:have_python

echo [2/4] ติดตั้งไลบรารี (ต้องต่ออินเทอร์เน็ต) ...
"%PY%" -m pip install --disable-pip-version-check -q openpyxl customtkinter tkinterdnd2
if errorlevel 1 (
  echo [ผิดพลาด] ติดตั้งไลบรารีไม่สำเร็จ - ตรวจอินเทอร์เน็ตแล้วลองใหม่
  goto :fail
)

echo [3/4] คัดลอกโปรแกรมไปที่ %APPDIR% ...
if not exist "%APPDIR%" mkdir "%APPDIR%"
copy /y "%SRC%\tongheer_app.pyw" "%APPDIR%\" >nul
copy /y "%SRC%\tongheer_invb.py" "%APPDIR%\" >nul
copy /y "%SRC%\app.ico" "%APPDIR%\" >nul
if exist "%SRC%\DATA_TONGHEER.xlsx" copy /y "%SRC%\DATA_TONGHEER.xlsx" "%APPDIR%\" >nul
rem ไม่เขียนทับค่าที่โปรแกรมจำไว้ ถ้าเคยติดตั้งแล้ว
if not exist "%APPDIR%\tongheer_map.json" if exist "%SRC%\tongheer_map.json" copy /y "%SRC%\tongheer_map.json" "%APPDIR%\" >nul

echo [4/4] สร้างไอคอนบน Desktop และ Start Menu ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$q=[char]34; $w=New-Object -ComObject WScript.Shell; foreach($f in @([Environment]::GetFolderPath('Desktop'),[Environment]::GetFolderPath('Programs'))){ $s=$w.CreateShortcut((Join-Path $f 'Tong Heer INVB.lnk')); $s.TargetPath=$env:PYW; $s.Arguments=$q+$env:APPDIR+'\tongheer_app.pyw'+$q; $s.WorkingDirectory=$env:APPDIR; $s.IconLocation=$env:APPDIR+'\app.ico'; $s.Description='Tong Heer INVB Converter'; $s.Save() }"
if errorlevel 1 (
  echo [ผิดพลาด] สร้างไอคอนไม่สำเร็จ - เปิดโปรแกรมได้จาก %APPDIR%\tongheer_app.pyw
  goto :fail
)

echo.
echo ติดตั้งเสร็จแล้ว - เปิดโปรแกรมจากไอคอน "Tong Heer INVB" บน Desktop
echo.
pause
exit /b 0

:fail
echo.
echo ติดตั้งไม่สำเร็จ
pause
exit /b 1
