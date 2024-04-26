@echo off
echo This script needs Administrator privilege.
echo To do so, right click on this script and select 'Run as administrator'.
echo Installing...
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Vietnamese-Font.ps1"
pause