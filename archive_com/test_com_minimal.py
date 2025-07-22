#!/usr/bin/env python3
"""
Minimal COM Connection Test
This script tests the most basic COM connection to Outlook.
"""

import sys

def test_basic_imports():
    """Test if we can import required modules"""
    print("Testing imports...")
    
    try:
        import win32com.client
        print("✅ win32com.client imported successfully")
    except ImportError as e:
        print(f"❌ Cannot import win32com.client: {e}")
        return False
    
    try:
        import pythoncom
        print("✅ pythoncom imported successfully")
    except ImportError as e:
        print(f"❌ Cannot import pythoncom: {e}")
        return False
    
    return True

def test_com_registration():
    """Test COM registration lookup"""
    print("\nTesting COM registration...")
    
    try:
        import pythoncom
        
        outlook_progids = [
            "Outlook.Application",
            "Outlook.Application.16",
            "Outlook.Application.15", 
            "Outlook.Application.14"
        ]
        
        found_any = False
        for progid in outlook_progids:
            try:
                clsid = pythoncom.CLSIDFromProgID(progid)
                print(f"✅ {progid} -> {clsid}")
                found_any = True
            except Exception as e:
                print(f"❌ {progid} -> {type(e).__name__}: {e}")
        
        return found_any
        
    except Exception as e:
        print(f"❌ COM registration test failed: {e}")
        return False

def test_dispatch():
    """Test basic Dispatch call"""
    print("\nTesting Dispatch calls...")
    
    try:
        import win32com.client as win32
        
        connection_methods = [
            ("Dispatch", lambda: win32.Dispatch("Outlook.Application")),
            ("DispatchEx", lambda: win32.DispatchEx("Outlook.Application")),
            ("GetActiveObject", lambda: win32.GetActiveObject("Outlook.Application")),
        ]
        
        for method_name, method_func in connection_methods:
            try:
                print(f"Trying {method_name}...")
                outlook = method_func()
                print(f"✅ {method_name} succeeded - got object: {type(outlook)}")
                
                # Try to get basic info
                try:
                    version = outlook.Version
                    print(f"✅ Outlook version: {version}")
                except:
                    print("⚠️ Could not get version")
                
                return True, outlook
                
            except Exception as e:
                print(f"❌ {method_name} failed: {type(e).__name__}: {e}")
        
        return False, None
        
    except Exception as e:
        print(f"❌ Dispatch test failed: {e}")
        return False, None

def test_namespace_access(outlook):
    """Test MAPI namespace access"""
    print("\nTesting MAPI namespace...")
    
    try:
        namespace = outlook.GetNamespace("MAPI")
        print("✅ Got MAPI namespace")
        
        try:
            # Try to get default folder (Inbox = 6)
            inbox = namespace.GetDefaultFolder(6)
            print("✅ Got default folder (Inbox)")
            
            try:
                folder_name = inbox.Name
                print(f"✅ Folder name: {folder_name}")
            except Exception as e:
                print(f"⚠️ Could not get folder name: {e}")
            
            try:
                item_count = inbox.Items.Count
                print(f"✅ Folder contains {item_count} items")
            except Exception as e:
                print(f"⚠️ Could not get item count: {e}")
            
            return True
            
        except Exception as e:
            print(f"❌ Could not access default folder: {type(e).__name__}: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Could not get MAPI namespace: {type(e).__name__}: {e}")
        return False

def test_user_permissions():
    """Test current user permissions"""
    print("\nTesting user permissions...")
    
    try:
        import ctypes
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        print(f"Running as Administrator: {is_admin}")
        
        if not is_admin:
            print("⚠️ Not running as Administrator - this may cause COM issues")
        
    except Exception as e:
        print(f"Could not check admin status: {e}")

def main():
    """Run all tests"""
    print("MINIMAL OUTLOOK COM CONNECTION TEST")
    print("="*50)
    
    # Test user permissions first
    test_user_permissions()
    
    # Test imports
    if not test_basic_imports():
        print("\n❌ Basic imports failed - install pywin32")
        return
    
    # Test COM registration
    if not test_com_registration():
        print("\n❌ COM registration failed - need to fix registration")
        return
    
    # Test dispatch
    success, outlook = test_dispatch()
    if not success:
        print("\n❌ Could not create Outlook object")
        print("\nNext steps:")
        print("1. Run as Administrator")
        print("2. Run fix_com_step_by_step.bat as Administrator")
        print("3. Use fix_com_permissions.py")
        return
    
    # Test namespace access
    if test_namespace_access(outlook):
        print("\n🎉 SUCCESS: Full COM connection working!")
        print("Your MailClerk Desktop handler should work now.")
    else:
        print("\n⚠️ PARTIAL SUCCESS: Can create Outlook object but cannot access MAPI")
        print("This suggests permission or security policy issues.")
    
    print("\n" + "="*50)
    print("Test completed.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nTest cancelled by user")
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        input("\nPress Enter to exit...")