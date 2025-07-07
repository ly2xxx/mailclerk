#!/usr/bin/env python3
"""
Test script to demonstrate dynamic handler discovery
"""

import os
import tempfile
import shutil
from email_handler_factory import get_handler_factory

def test_handler_discovery():
    """Test that handlers are discovered dynamically"""
    
    print("DYNAMIC EMAIL HANDLER DISCOVERY TEST")
    print("="*50)
    
    # Get factory and show current handlers
    factory = get_handler_factory()
    available = factory.get_available_handlers()
    
    print(f"Found {len(available)} handlers:")
    for key, info in available.items():
        print(f"  {key}: {info['name']}")
        print(f"    Module: {info['module']}")
        print(f"    Class: {info['class_name']}")
        print()
    
    # Test creating handlers
    print("Testing handler creation:")
    for key in available.keys():
        try:
            handler = factory.create_handler(key)
            if handler:
                print(f"  ✅ {key}: Created successfully")
                handler.close_connection()  # Clean up
            else:
                print(f"  ❌ {key}: Creation failed")
        except Exception as e:
            print(f"  ❌ {key}: Error - {e}")
    
    print("\nTesting refresh functionality:")
    factory.refresh_handlers()
    new_available = factory.get_available_handlers()
    print(f"After refresh: {len(new_available)} handlers")
    
    return len(available) > 0

def test_handler_removal_simulation():
    """Simulate what happens when a handler module is unavailable"""
    
    print("\nHANDLER REMOVAL SIMULATION")
    print("="*50)
    
    # Create a backup directory
    backup_dir = tempfile.mkdtemp()
    
    try:
        # Move email_handler.py to backup (simulate removal)
        if os.path.exists('email_handler.py'):
            shutil.move('email_handler.py', os.path.join(backup_dir, 'email_handler.py'))
            print("Moved email_handler.py (simulating removal)")
        
        # Create new factory and check handlers
        from importlib import reload
        import email_handler_factory
        reload(email_handler_factory)
        
        new_factory = email_handler_factory.EmailHandlerFactory()
        available = new_factory.get_available_handlers()
        
        print(f"After 'removing' desktop handler: {len(available)} handlers")
        for key, info in available.items():
            print(f"  {key}: {info['name']}")
        
        # Restore the file
        backup_file = os.path.join(backup_dir, 'email_handler.py')
        if os.path.exists(backup_file):
            shutil.move(backup_file, 'email_handler.py')
            print("Restored email_handler.py")
        
        # Test again after restore
        reload(email_handler_factory)
        restored_factory = email_handler_factory.EmailHandlerFactory()
        restored_available = restored_factory.get_available_handlers()
        
        print(f"After restoring: {len(restored_available)} handlers")
        for key, info in restored_available.items():
            print(f"  {key}: {info['name']}")
    
    finally:
        # Clean up backup directory
        shutil.rmtree(backup_dir, ignore_errors=True)

def test_platform_filtering():
    """Test platform-specific handler filtering"""
    
    print("\nPLATFORM FILTERING TEST")
    print("="*50)
    
    import platform
    current_platform = platform.system()
    print(f"Current platform: {current_platform}")
    
    factory = get_handler_factory()
    available = factory.get_available_handlers()
    
    print("Available handlers for this platform:")
    for key, info in available.items():
        print(f"  {key}: {info['name']}")
    
    if current_platform != 'Windows':
        desktop_available = 'desktop' in available
        print(f"Desktop handler available on {current_platform}: {desktop_available}")
        if not desktop_available:
            print("  ✅ Correctly filtered out Windows-only handler")
    else:
        print("  Running on Windows - all handlers should be available")

if __name__ == "__main__":
    try:
        test_handler_discovery()
        test_platform_filtering()
        # Uncomment to test removal simulation (be careful!)
        # test_handler_removal_simulation()
        
        print("\n🎉 All tests completed!")
        
    except Exception as e:
        print(f"\n❌ Test failed: {e}")
        import traceback
        traceback.print_exc()