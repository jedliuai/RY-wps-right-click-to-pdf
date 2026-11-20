@echo off
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup.ps1" -Uninstall
set "setupExit=%errorlevel%"
if not "%setupExit%"=="0" pause
exit /b %setupExit%
