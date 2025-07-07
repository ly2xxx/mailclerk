#!/usr/bin/env python3
"""
Advanced COM Permission Fix for Outlook
This script attempts to fix COM registration and permission issues.
"""

import os
import sys
import subprocess
import winreg
import logging
import ctypes
from pathlib import Path

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def is_admin():
    """Check if running as administrator"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """Restart script as administrator"""
    if is_admin():
        return True
    else:
        logger.warning("This script requires Administrator privileges")
        logger.info("Attempting to restart as Administrator...")
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
            return False
        except Exception as e:
            logger.error(f"Failed to run as admin: {e}")
            return False

def find_outlook_path():
    """Find Outlook installation path"""
    possible_paths = [
        r"C:\Program Files\Microsoft Office\root\Office16",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16", 
        r"C:\Program Files\Microsoft Office\Office16",
        r"C:\Program Files (x86)\Microsoft Office\Office16",
        r"C:\Program Files\Microsoft Office\Office15",
        r"C:\Program Files (x86)\Microsoft Office\Office15",
        r"C:\Program Files\Microsoft Office\Office14",
        r"C:\Program Files (x86)\Microsoft Office\Office14",
    ]
    
    for path in possible_paths:
        outlook_exe = Path(path) / "OUTLOOK.EXE"
        if outlook_exe.exists():
            logger.info(f"Found Outlook at: {path}")
            return path
    
    logger.error("Could not find Outlook installation")
    return None

def kill_outlook():
    """Kill all Outlook processes"""
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'OUTLOOK.EXE', '/T'], 
                      capture_output=True, check=False)
        logger.info("Outlook processes terminated")
        return True
    except Exception as e:
        logger.warning(f"Could not kill Outlook processes: {e}")
        return False

def register_dll(dll_path, use_install_flag=False):
    """Register a DLL with regsvr32"""
    try:
        if not Path(dll_path).exists():
            logger.warning(f"DLL not found: {dll_path}")
            return False
        
        cmd = ['regsvr32', '/s']
        if use_install_flag:
            cmd.extend(['/i', '/n'])
        cmd.append(str(dll_path))
        
        result = subprocess.run(cmd, capture_output=True)
        
        if result.returncode == 0:
            logger.info(f"✅ Successfully registered: {Path(dll_path).name}")
            return True
        else:
            logger.error(f"❌ Failed to register: {Path(dll_path).name}")
            return False
            
    except Exception as e:
        logger.error(f"Error registering {dll_path}: {e}")
        return False

def fix_registry_permissions():
    """Fix registry permissions for COM classes"""
    try:
        # Registry keys that need to be accessible
        registry_keys = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\Outlook.Application"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\CLSID\{0006F03A-0000-0000-C000-000000000046}"),
            (winreg.HKEY_CLASSES_ROOT, r"Outlook.Application"),
            (winreg.HKEY_CLASSES_ROOT, r"CLSID\{0006F03A-0000-0000-C000-000000000046}"),
        ]
        
        for hive, key_path in registry_keys:
            try:
                # Try to open/create the key
                with winreg.CreateKey(hive, key_path) as key:
                    logger.info(f"✅ Registry key accessible: {key_path}")
            except Exception as e:
                logger.warning(f"Registry key issue: {key_path} - {e}")
        
        return True
        
    except Exception as e:
        logger.error(f"Registry permission fix failed: {e}")
        return False

def clear_outlook_security():
    """Clear Outlook security restrictions"""
    try:
        security_keys = [
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Outlook\Security"),
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\15.0\Outlook\Security"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\16.0\Outlook\Security"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\15.0\Outlook\Security"),
        ]
        
        for hive, key_path in security_keys:
            try:
                winreg.DeleteKey(hive, key_path)
                logger.info(f"Cleared security key: {key_path}")
            except FileNotFoundError:
                pass  # Key doesn't exist, that's fine
            except Exception as e:
                logger.warning(f"Could not clear security key {key_path}: {e}")
        
        return True
        
    except Exception as e:
        logger.error(f"Security clearing failed: {e}")
        return False

def restart_services():
    """Restart Windows services that might help"""
    services = ['winmgmt']  # Windows Management Instrumentation
    
    for service in services:
        try:
            # Stop service
            subprocess.run(['net', 'stop', service, '/y'], 
                          capture_output=True, check=False)
            
            # Start service
            result = subprocess.run(['net', 'start', service], 
                                  capture_output=True, check=False)
            
            if result.returncode == 0:
                logger.info(f"✅ Restarted service: {service}")
            else:
                logger.warning(f"Could not restart service: {service}")
                
        except Exception as e:
            logger.warning(f"Error with service {service}: {e}")

def test_com_connection():
    """Test COM connection to Outlook"""
    try:
        import win32com.client as win32
        
        logger.info("Testing COM connection...")
        outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        folder_name = inbox.Name
        item_count = inbox.Items.Count
        
        logger.info(f"🎉 SUCCESS: Connected to '{folder_name}' with {item_count} items")
        return True
        
    except Exception as e:
        logger.error(f"❌ COM connection failed: {e}")
        return False

def main():
    """Main fix routine"""
    print("ADVANCED OUTLOOK COM PERMISSION FIX")
    print("="*50)
    
    # Check admin privileges
    if not run_as_admin():
        return
    
    logger.info("Running as Administrator ✅")
    
    # Find Outlook
    outlook_path = find_outlook_path()
    if not outlook_path:
        logger.error("Cannot proceed without Outlook installation")
        return
    
    # Step-by-step fixes
    fixes = [
        ("Killing Outlook processes", kill_outlook),
        ("Fixing registry permissions", fix_registry_permissions),
        ("Clearing security restrictions", clear_outlook_security),
        ("Restarting Windows services", restart_services),
    ]
    
    # Register DLLs
    dlls_to_register = [
        ("outlctl.dll", False),
        ("msoutl.olb", False), 
        ("olmapi32.dll", False),
        ("outllib.dll", False),
        ("outlctl.dll", True),  # Register again with /i flag
    ]
    
    # Run general fixes
    for fix_name, fix_func in fixes:
        logger.info(f"\n--- {fix_name} ---")
        try:
            success = fix_func()
            if success:
                logger.info(f"✅ {fix_name} completed")
            else:
                logger.warning(f"⚠️ {fix_name} had issues")
        except Exception as e:
            logger.error(f"❌ {fix_name} failed: {e}")
    
    # Register DLLs
    logger.info(f"\n--- Registering COM components ---")
    for dll_name, use_flag in dlls_to_register:
        dll_path = Path(outlook_path) / dll_name
        register_dll(dll_path, use_flag)
    
    # Start Outlook
    logger.info(f"\n--- Starting Outlook ---")
    try:
        outlook_exe = Path(outlook_path) / "OUTLOOK.EXE"
        subprocess.Popen([str(outlook_exe)])
        logger.info("Outlook started, waiting 10 seconds...")
        
        import time
        time.sleep(10)
        
    except Exception as e:
        logger.error(f"Could not start Outlook: {e}")
    
    # Test connection
    logger.info(f"\n--- Testing COM Connection ---")
    if test_com_connection():
        print("\n🎉 SUCCESS: Outlook COM connection is now working!")
        print("You can now use the Desktop handler in MailClerk.")
    else:
        print("\n❌ COM connection still fails.")
        print("Recommended next steps:")
        print("1. Restart your computer and try again")
        print("2. Repair Office installation via Control Panel")
        print("3. Use the Browser handler as a reliable alternative")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        input("\nPress Enter to exit...")