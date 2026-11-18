@echo off
setlocal
cd /d "%~dp0"
if not exist "Logs" mkdir "Logs"
%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Banmedica.ps1" 1>>"%~dp0Logs\inicio.log" 2>>&1
if errorlevel 1 (
 echo La aplicacion encontro un error. Revise Logs\inicio.log
 pause
)
endlocal
