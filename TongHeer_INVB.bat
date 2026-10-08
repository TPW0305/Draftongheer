@echo off
rem Double-click to open, or drag an Export Summary Report (.xlsx) onto this file.
set "PYW=%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe"
if not exist "%PYW%" set "PYW=pythonw"
start "" "%PYW%" "%~dp0tongheer_app.pyw" %*
