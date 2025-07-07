#!/usr/bin/env python3
"""
Outlook COM Connection Diagnostic Tool
This script helps diagnose common Outlook connection issues.
"""

import sys
import platform
import subprocess
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def check_windows():
    """Check if running on Windows"""
    if platform.system() != 'Windows':
        logger.error("This diagnostic tool only works on Windows")
        return False
    
    logger.info(f"Running on {platform.system()} {platform.release()} ({platform.machine()})")
    return True

def check_python_architecture():
    """Check Python architecture (32-bit vs 64-bit)"""
    arch = platform.architecture()[0]
    logger.info(f"Python architecture: {arch}")
    
    if arch == '32bit':
        logger.warning("Using 32-bit Python. Ensure you have 32-bit Office installed or use 64-bit Python with 64-bit Office.")
    
    return arch

def check_office_installation():
    """Check for Office/Outlook installation in registry"""
    try:
        import winreg
        
        office_versions = {
            "16.0": "Office 2016/2019/2021/365",
            "15.0": "Office 2013", 
            "14.0": "Office 2010",
            "12.0": "Office 2007"
        }
        
        found_versions = []
        
        for version, name in office_versions.items():
            try:
                key_path = f"SOFTWARE\\Microsoft\\Office\\{version}\\Outlook"
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path):
                    found_versions.append(f"{name} (version {version})")
                    logger.info(f"Found: {name}")
            except FileNotFoundError:
                continue
            except Exception as e:
                logger.warning(f"Error checking {name}: {e}")
        
        if not found_versions:
            logger.error("No Office/Outlook installation found in registry")
            return False
        
        logger.info(f"Detected Office versions: {', '.join(found_versions)}")
        return True
        
    except ImportError:
        logger.error("Cannot access Windows registry (winreg module not available)")
        return False
    except Exception as e:
        logger.error(f"Registry check failed: {e}")
        return False

def check_outlook_process():
    """Check if Outlook is running"""
    try:
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq OUTLOOK.EXE'],
            capture_output=True, text=True, timeout=10
        )
        
        if 'OUTLOOK.EXE' in result.stdout:
            logger.info("Outlook process is running")
            return True
        else:
            logger.warning("Outlook process is not running")
            return False
            
    except subprocess.TimeoutExpired:
        logger.error("Timeout checking for Outlook process")
        return False
    except Exception as e:
        logger.error(f"Error checking Outlook process: {e}")
        return False

def check_pywin32():
    """Check pywin32 installation"""
    try:
        import win32com.client
        logger.info("pywin32 is installed and importable")
        return True
    except ImportError as e:
        logger.error(f"pywin32 not available: {e}")
        logger.error("Install with: pip install pywin32")
        return False
    except Exception as e:
        logger.error(f"Error importing pywin32: {e}")
        return False

def check_com_registration():
    """Check COM registration for Outlook"""
    try:
        import pythoncom
        
        outlook_progids = [
            "Outlook.Application",
            "Outlook.Application.16",
            "Outlook.Application.15", 
            "Outlook.Application.14"
        ]
        
        working_progids = []
        
        for progid in outlook_progids:
            try:
                clsid = pythoncom.CLSIDFromProgID(progid)
                working_progids.append(progid)
                logger.info(f"COM registration OK for {progid} -> {clsid}")
            except Exception:
                logger.warning(f"COM registration failed for {progid}")
        
        if working_progids:
            logger.info(f"Working COM registrations: {', '.join(working_progids)}")
            return True
        else:
            logger.error("No working COM registrations found")
            return False
            
    except ImportError:
        logger.error("pythoncom not available")
        return False
    except Exception as e:
        logger.error(f"COM registration check failed: {e}")
        return False

def test_outlook_connection():
    """Test actual Outlook connection"""
    try:
        import win32com.client as win32
        
        connection_methods = [
            "Outlook.Application",
            "Outlook.Application.16",
            "Outlook.Application.15", 
            "Outlook.Application.14"
        ]
        
        for method in connection_methods:
            try:
                logger.info(f"Testing connection with {method}...")
                outlook = win32.Dispatch(method)
                namespace = outlook.GetNamespace("MAPI")
                inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
                folder_name = inbox.Name
                
                logger.info(f"✅ SUCCESS: Connected to '{folder_name}' using {method}")
                return True
                
            except Exception as e:
                logger.warning(f"❌ FAILED: {method} -> {type(e).__name__}: {e}")
                continue
        
        logger.error("All connection methods failed")
        return False
        
    except Exception as e:
        logger.error(f"Connection test failed: {e}")
        return False

def print_solutions():
    """Print common solutions"""
    print("\n" + "="*60)
    print("COMMON SOLUTIONS FOR COM_ERROR(-2147221005)")
    print("="*60)
    
    solutions = [
        "1. RESTART OUTLOOK",
        "   - Close Outlook completely",
        "   - Start Outlook and let it fully load",
        "   - Try the connection again",
        "",
        "2. RUN AS ADMINISTRATOR",
        "   - Right-click your terminal/IDE",
        "   - Select 'Run as Administrator'",
        "   - Try running the application again",
        "",
        "3. REPAIR OFFICE INSTALLATION",
        "   - Go to Control Panel > Programs",
        "   - Find Microsoft Office",
        "   - Click 'Change' > 'Quick Repair'",
        "",
        "4. RE-REGISTER COM COMPONENTS",
        "   - Open Command Prompt as Administrator",
        "   - Run: regsvr32 /i /n /s outlctl.dll",
        "   - Run: regsvr32 /s msoutl.olb",
        "",
        "5. CHECK ARCHITECTURE MISMATCH",
        "   - Ensure Python and Office have same architecture",
        "   - 32-bit Python needs 32-bit Office",
        "   - 64-bit Python needs 64-bit Office",
        "",
        "6. REINSTALL PYWIN32",
        "   - pip uninstall pywin32",
        "   - pip install pywin32",
        "   - python Scripts/pywin32_postinstall.py -install",
        "",
        "7. CHECK WINDOWS PERMISSIONS",
        "   - Ensure user has permission to access MAPI",
        "   - Try different user account if needed"
    ]
    
    for solution in solutions:
        print(solution)

def main():
    """Run all diagnostic checks"""
    print("OUTLOOK COM CONNECTION DIAGNOSTIC TOOL")
    print("="*50)
    
    checks = [
        ("Windows OS", check_windows),
        ("Python Architecture", check_python_architecture), 
        ("Office Installation", check_office_installation),
        ("Outlook Process", check_outlook_process),
        ("pywin32 Module", check_pywin32),
        ("COM Registration", check_com_registration),
        ("Outlook Connection", test_outlook_connection)
    ]
    
    results = {}
    
    for check_name, check_func in checks:
        print(f"\n--- {check_name} ---")
        try:
            result = check_func()
            results[check_name] = result
        except Exception as e:
            logger.error(f"Check '{check_name}' crashed: {e}")
            results[check_name] = False
    
    # Summary
    print(f"\n{'='*50}")
    print("DIAGNOSTIC SUMMARY")
    print(f"{'='*50}")
    
    for check_name, result in results.items():
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{check_name:20} : {status}")
    
    # Show solutions if any checks failed
    if not all(results.values()):
        print_solutions()
    else:
        print("\n🎉 All checks passed! Outlook connection should work.")

if __name__ == "__main__":
    main()