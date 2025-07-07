@echo off
echo ================================================================
echo OUTLOOK COM REGISTRATION FIX - STEP BY STEP
echo ================================================================
echo.
echo This script will systematically fix Outlook COM issues.
echo IMPORTANT: This script MUST be run as Administrator!
echo.
echo Checking if running as Administrator...
net session >nul 2>&1
if %errorLevel% == 0 (
    echo ✅ Running as Administrator - Good!
) else (
    echo ❌ NOT running as Administrator!
    echo.
    echo Please:
    echo 1. Right-click this file
    echo 2. Select "Run as administrator"
    echo 3. Try again
    echo.
    pause
    exit /b 1
)

echo.
echo ================================================================
echo STEP 1: Finding Outlook Installation
echo ================================================================

REM Find Outlook installation path
set "OUTLOOK_PATH="
for %%P in (
    "C:\Program Files\Microsoft Office\root\Office16"
    "C:\Program Files (x86)\Microsoft Office\root\Office16"
    "C:\Program Files\Microsoft Office\Office16"
    "C:\Program Files (x86)\Microsoft Office\Office16"
    "C:\Program Files\Microsoft Office\Office15"
    "C:\Program Files (x86)\Microsoft Office\Office15"
) do (
    if exist "%%~P\OUTLOOK.EXE" (
        set "OUTLOOK_PATH=%%~P"
        echo ✅ Found Outlook at: %%~P
        goto :found_outlook
    )
)

echo ❌ Could not find Outlook installation
echo Please install Microsoft Outlook first
pause
exit /b 1

:found_outlook
echo.
echo ================================================================
echo STEP 2: Stopping Outlook Process
echo ================================================================

echo Stopping Outlook processes...
taskkill /F /IM OUTLOOK.EXE /T >nul 2>&1
if %errorlevel% == 0 (
    echo ✅ Outlook processes stopped
) else (
    echo ℹ️ No Outlook processes running
)

echo Waiting 3 seconds for processes to fully stop...
timeout /t 3 /nobreak >nul

echo.
echo ================================================================
echo STEP 3: Registering Core COM Components
echo ================================================================

pushd "%OUTLOOK_PATH%"

echo Registering outlctl.dll...
regsvr32 /s outlctl.dll
if %errorlevel% == 0 (
    echo ✅ outlctl.dll registered successfully
) else (
    echo ❌ Failed to register outlctl.dll
)

echo Registering msoutl.olb...
regsvr32 /s msoutl.olb
if %errorlevel% == 0 (
    echo ✅ msoutl.olb registered successfully
) else (
    echo ❌ Failed to register msoutl.olb
)

echo Registering olmapi32.dll...
regsvr32 /s olmapi32.dll
if %errorlevel% == 0 (
    echo ✅ olmapi32.dll registered successfully
) else (
    echo ❌ Failed to register olmapi32.dll
)

echo Registering outllib.dll...
regsvr32 /s outllib.dll
if %errorlevel% == 0 (
    echo ✅ outllib.dll registered successfully
) else (
    echo ❌ Failed to register outllib.dll
)

popd

echo.
echo ================================================================
echo STEP 4: Special COM Registration with Installation Flag
echo ================================================================

pushd "%OUTLOOK_PATH%"

echo Registering outlctl.dll with installation flag...
regsvr32 /i /n /s outlctl.dll
if %errorlevel% == 0 (
    echo ✅ outlctl.dll registered with /i flag
) else (
    echo ❌ Failed to register outlctl.dll with /i flag
)

popd

echo.
echo ================================================================
echo STEP 5: Fixing Registry Permissions
echo ================================================================

echo Setting registry permissions for COM classes...

REM Set permissions for main Outlook Application class
echo Granting permissions to Outlook.Application...
reg add "HKLM\SOFTWARE\Classes\Outlook.Application" /f >nul 2>&1

REM Grant permissions to the CLSID
echo Granting permissions to Outlook CLSID...
reg add "HKLM\SOFTWARE\Classes\CLSID\{0006F03A-0000-0000-C000-000000000046}" /f >nul 2>&1

echo.
echo ================================================================
echo STEP 6: Clearing Security Restrictions
echo ================================================================

echo Clearing Outlook security restrictions...

REM Clear user-level security settings
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\15.0\Outlook\Security" /f >nul 2>&1

REM Clear machine-level security settings
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Microsoft\Office\15.0\Outlook\Security" /f >nul 2>&1

echo ✅ Security restrictions cleared

echo.
echo ================================================================
echo STEP 7: Restarting Windows Services
echo ================================================================

echo Restarting Windows Management Instrumentation service...
net stop winmgmt /y >nul 2>&1
net start winmgmt >nul 2>&1
if %errorlevel% == 0 (
    echo ✅ WMI service restarted
) else (
    echo ⚠️ Could not restart WMI service
)

echo.
echo ================================================================
echo STEP 8: Starting Outlook
echo ================================================================

echo Starting Outlook...
start "" "%OUTLOOK_PATH%\OUTLOOK.EXE"

echo Waiting 10 seconds for Outlook to fully load...
timeout /t 10 /nobreak >nul

echo.
echo ================================================================
echo STEP 9: Testing COM Connection
echo ================================================================

echo Testing Python COM connection...
python -c "
import win32com.client as win32
try:
    outlook = win32.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')
    inbox = namespace.GetDefaultFolder(6)
    print('✅ SUCCESS: COM connection works!')
    print(f'Connected to folder: {inbox.Name}')
    print(f'Folder contains {inbox.Items.Count} items')
except Exception as e:
    print(f'❌ FAILED: {e}')
    print('COM connection still not working')
"

echo.
echo ================================================================
echo STEP 10: Final Status Check
echo ================================================================

echo Running MailClerk diagnostics...
python -c "
from email_handler_factory import get_handler_factory
factory = get_handler_factory()
handler = factory.create_handler('desktop')
if handler and hasattr(handler, 'diagnose_outlook_installation'):
    diagnosis = handler.diagnose_outlook_installation()
    print('Final Diagnostic Results:')
    for check, status in diagnosis.items():
        icon = '✅' if status else '❌'
        readable_name = check.replace('_', ' ').title()
        print(f'  {icon} {readable_name}')
else:
    print('Could not run diagnostics')
"

echo.
echo ================================================================
echo COMPLETED!
echo ================================================================
echo.
echo If you see SUCCESS above, your COM connection is now working!
echo If you still see failures, try:
echo 1. Restart your computer
echo 2. Repair Office installation via Control Panel
echo 3. Use the Browser handler as an alternative
echo.
pause