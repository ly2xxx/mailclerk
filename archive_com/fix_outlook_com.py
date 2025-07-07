#!/usr/bin/env python3
"""
Outlook COM Fix Tool
This script attempts to fix common Outlook COM registration issues.
"""

import os
import sys
import subprocess
import platform
import logging
from pathlib import Path

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def is_admin():
    """Check if running as administrator"""
    try:
        return os.getuid() == 0
    except AttributeError:
        # Windows
        try:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

def run_as_admin():
    """Rerun the script as administrator"""
    if platform.system() == 'Windows':
        try:
            import ctypes
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, 
                " ".join(sys.argv), None, 1
            )
        except Exception as e:
            logger.error(f"Failed to run as admin: {e}")
            logger.info("Please manually run this script as Administrator")

def find_office_path():
    """Find Office installation path"""
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
    
    logger.warning("Could not find Outlook installation path")
    return None

def register_com_components():
    """Register Outlook COM components"""
    office_path = find_office_path()
    if not office_path:
        return False
    
    # DLLs to register
    dlls_to_register = [
        "outlctl.dll",
        "msoutl.olb", 
        "olmapi32.dll",
        "outllib.dll"
    ]
    
    success_count = 0
    
    for dll in dlls_to_register:
        dll_path = Path(office_path) / dll
        if dll_path.exists():
            try:
                cmd = f'regsvr32 /s "{dll_path}"'
                logger.info(f"Registering: {dll}")
                result = subprocess.run(cmd, shell=True, capture_output=True)
                if result.returncode == 0:
                    logger.info(f"✅ Successfully registered {dll}")
                    success_count += 1
                else:
                    logger.warning(f"❌ Failed to register {dll}")
            except Exception as e:
                logger.error(f"Error registering {dll}: {e}")
        else:
            logger.warning(f"DLL not found: {dll_path}")
    
    return success_count > 0

def fix_registry_permissions():
    """Fix registry permissions for COM registration"""
    try:
        # Grant permissions to HKEY_CLASSES_ROOT
        registry_keys = [
            r"HKEY_CLASSES_ROOT\Outlook.Application",
            r"HKEY_CLASSES_ROOT\CLSID\{0006F03A-0000-0000-C000-000000000046}",
            r"HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Outlook.Application"
        ]
        
        for key in registry_keys:
            try:
                # Use icacls to grant permissions
                cmd = f'icacls "{key}" /grant Everyone:F /t'
                subprocess.run(cmd, shell=True, capture_output=True)
                logger.info(f"Set permissions for: {key}")
            except Exception as e:
                logger.warning(f"Could not set permissions for {key}: {e}")
        
        return True
    except Exception as e:
        logger.error(f"Registry permission fix failed: {e}")
        return False

def kill_outlook_processes():
    """Kill all Outlook processes"""
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'OUTLOOK.EXE'], 
                      capture_output=True, check=False)
        logger.info("Killed Outlook processes")
        return True
    except Exception as e:
        logger.warning(f"Could not kill Outlook processes: {e}")
        return False

def start_outlook():
    """Start Outlook"""
    office_path = find_office_path()
    if office_path:
        try:
            outlook_exe = Path(office_path) / "OUTLOOK.EXE"
            subprocess.Popen([str(outlook_exe)], shell=True)
            logger.info("Started Outlook")
            return True
        except Exception as e:
            logger.error(f"Could not start Outlook: {e}")
    return False

def reinstall_pywin32():
    """Reinstall pywin32 package"""
    try:
        logger.info("Reinstalling pywin32...")
        
        # Uninstall
        subprocess.run([sys.executable, '-m', 'pip', 'uninstall', 'pywin32', '-y'], 
                      capture_output=True)
        
        # Reinstall
        result = subprocess.run([sys.executable, '-m', 'pip', 'install', 'pywin32'], 
                               capture_output=True, text=True)
        
        if result.returncode == 0:
            logger.info("✅ pywin32 reinstalled successfully")
            
            # Run post-install script
            try:
                import site
                scripts_path = Path(site.getsitepackages()[0]).parent / "Scripts"
                postinstall_script = scripts_path / "pywin32_postinstall.py"
                
                if postinstall_script.exists():
                    subprocess.run([sys.executable, str(postinstall_script), '-install'],
                                 capture_output=True)
                    logger.info("✅ pywin32 post-install completed")
                else:
                    logger.warning("Could not find pywin32_postinstall.py")
            except Exception as e:
                logger.warning(f"Post-install script failed: {e}")
            
            return True
        else:
            logger.error(f"pywin32 reinstall failed: {result.stderr}")
            return False
            
    except Exception as e:
        logger.error(f"Error reinstalling pywin32: {e}")
        return False

def test_connection():
    """Test Outlook connection after fixes"""
    try:
        import win32com.client as win32
        
        logger.info("Testing Outlook connection...")
        outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        folder_name = inbox.Name
        
        logger.info(f"🎉 SUCCESS: Connected to '{folder_name}'")
        return True
        
    except Exception as e:
        logger.error(f"❌ Connection test still fails: {e}")
        return False

def main():
    """Main fix routine"""
    print("OUTLOOK COM FIX TOOL")
    print("="*40)
    
    if platform.system() != 'Windows':
        logger.error("This tool only works on Windows")
        return
    
    if not is_admin():
        logger.warning("Not running as Administrator")
        logger.info("Some fixes require Administrator privileges")
        logger.info("Attempting to restart as Administrator...")
        run_as_admin()
        return
    
    logger.info("Running as Administrator ✅")
    
    fixes = [
        ("Killing Outlook processes", kill_outlook_processes),
        ("Registering COM components", register_com_components),
        ("Fixing registry permissions", fix_registry_permissions),
        ("Reinstalling pywin32", reinstall_pywin32),
    ]
    
    for fix_name, fix_func in fixes:
        print(f"\n--- {fix_name} ---")
        try:
            success = fix_func()
            if success:
                logger.info(f"✅ {fix_name} completed")
            else:
                logger.warning(f"⚠️ {fix_name} had issues")
        except Exception as e:
            logger.error(f"❌ {fix_name} failed: {e}")
    
    print(f"\n{'='*40}")
    print("TESTING CONNECTION")
    print(f"{'='*40}")
    
    # Wait a moment for things to settle
    import time
    time.sleep(2)
    
    if test_connection():
        print("\n🎉 SUCCESS: Outlook COM connection is now working!")
        print("You can now run your MailClerk application.")
    else:
        print("\n❌ Connection still fails. Try these manual steps:")
        print("1. Restart your computer")
        print("2. Repair Office installation via Control Panel")
        print("3. Check for Windows updates")
        print("4. Reinstall Microsoft Office")

if __name__ == "__main__":
    main()