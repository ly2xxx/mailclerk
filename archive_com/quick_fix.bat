@echo off
echo OUTLOOK COM QUICK FIX
echo ====================
echo.
echo This script will attempt to fix Outlook COM registration issues.
echo You may need to run this as Administrator.
echo.
pause

echo Killing Outlook processes...
taskkill /F /IM OUTLOOK.EXE >nul 2>&1

echo.
echo Registering COM components...
regsvr32 /s outlctl.dll
regsvr32 /s msoutl.olb
regsvr32 /s olmapi32.dll
regsvr32 /s outllib.dll

echo.
echo Registering with /i flag...
regsvr32 /i /n /s outlctl.dll

echo.
echo Starting Outlook...
start outlook

echo.
echo Waiting for Outlook to fully load...
timeout /t 10 /nobreak >nul

echo.
echo Testing Python connection...
python -c "import win32com.client; o=win32com.client.Dispatch('Outlook.Application'); print('SUCCESS: Connected to Outlook')"

echo.
echo Fix attempt completed.
echo If you still get errors, try:
echo 1. Run this script as Administrator
echo 2. Restart your computer
echo 3. Repair Office installation
echo.
pause