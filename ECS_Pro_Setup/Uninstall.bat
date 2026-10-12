@echo off
chcp 65001 >nul
title ถอนการติดตั้ง ECS Pro
set "APPDIR=%LOCALAPPDATA%\TongHeerINVB"
echo จะลบโปรแกรม ECS Pro และค่าที่โปรแกรมจำไว้ (tongheer_map.json) ออกจากเครื่องนี้
echo Python จะไม่ถูกลบ
choice /c YN /m "ยืนยันการลบ"
if errorlevel 2 exit /b 0
powershell -NoProfile -Command "foreach($f in @([Environment]::GetFolderPath('Desktop'),[Environment]::GetFolderPath('Programs'))){ foreach($n in @('ECS Pro.lnk','Tong Heer INVB.lnk')){ $p=Join-Path $f $n; if(Test-Path -LiteralPath $p){ Remove-Item -LiteralPath $p -Force } } }"
if exist "%APPDIR%\tongheer_app.pyw" rmdir /s /q "%APPDIR%"
echo ลบเรียบร้อย
pause
