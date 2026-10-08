@echo off
chcp 65001 >nul
title ถอนการติดตั้ง Tong Heer INVB
set "APPDIR=%LOCALAPPDATA%\TongHeerINVB"
echo จะลบโปรแกรม Tong Heer INVB และค่าที่โปรแกรมจำไว้ (tongheer_map.json) ออกจากเครื่องนี้
echo Python จะไม่ถูกลบ
choice /c YN /m "ยืนยันการลบ"
if errorlevel 2 exit /b 0
powershell -NoProfile -Command "foreach($f in @([Environment]::GetFolderPath('Desktop'),[Environment]::GetFolderPath('Programs'))){ $p=Join-Path $f 'Tong Heer INVB.lnk'; if(Test-Path -LiteralPath $p){ Remove-Item -LiteralPath $p -Force } }"
if exist "%APPDIR%\tongheer_app.pyw" rmdir /s /q "%APPDIR%"
echo ลบเรียบร้อย
pause
