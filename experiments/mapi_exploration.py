"""
MAPI Exploration - Milestone 1
==============================
Goal: Discover what's available in pywin32's MAPI bindings

Run this on Windows with Outlook installed:
    python experiments/mapi_exploration.py

This will tell us:
1. What modules are available in win32com.mapi
2. What functions/classes are exposed
3. If basic MAPI initialization works
4. What constants (property tags) are defined
"""

import sys

def print_section(title):
    print(f"\n{'='*60}")
    print(f" {title}")
    print('='*60)

def main():
    print("MAPI Exploration Script - Milestone 1")
    print(f"Python: {sys.version}")
    print(f"Platform: {sys.platform}")
    
    if sys.platform != 'win32':
        print("\n❌ ERROR: This script must be run on Windows!")
        return
    
    # =========================================================================
    # Step 1: Check if pywin32 is installed
    # =========================================================================
    print_section("Step 1: Checking pywin32 installation")
    
    try:
        import win32com
        print(f"✓ win32com found at: {win32com.__file__}")
    except ImportError as e:
        print(f"❌ win32com not installed: {e}")
        print("   Install with: pip install pywin32")
        return
    
    try:
        import pythoncom
        print(f"✓ pythoncom found")
    except ImportError as e:
        print(f"❌ pythoncom not found: {e}")
        return
    
    # =========================================================================
    # Step 2: Explore win32com.mapi module
    # =========================================================================
    print_section("Step 2: Exploring win32com.mapi")
    
    try:
        from win32com import mapi
        print(f"✓ win32com.mapi found at: {mapi.__file__}")
        
        # List all attributes
        mapi_attrs = [a for a in dir(mapi) if not a.startswith('_')]
        print(f"\nAvailable in win32com.mapi ({len(mapi_attrs)} items):")
        for attr in sorted(mapi_attrs):
            obj = getattr(mapi, attr)
            obj_type = type(obj).__name__
            print(f"  {attr}: {obj_type}")
            
    except ImportError as e:
        print(f"❌ win32com.mapi not available: {e}")
        print("   This might mean MAPI bindings aren't included")
    
    # =========================================================================
    # Step 3: Check for mapitags (property tag constants)
    # =========================================================================
    print_section("Step 3: Checking mapitags (property constants)")
    
    try:
        from win32com.mapi import mapitags
        print(f"✓ mapitags found")
        
        # Look for date-related property tags
        date_tags = [a for a in dir(mapitags) if 'TIME' in a or 'DATE' in a or 'SUBMIT' in a or 'DELIVERY' in a]
        print(f"\nDate-related property tags ({len(date_tags)} found):")
        for tag in sorted(date_tags):
            value = getattr(mapitags, tag)
            print(f"  {tag} = 0x{value:08X}")
        
        # Check for specific tags we need
        print("\n*** KEY TAGS WE NEED ***")
        key_tags = [
            'PR_MESSAGE_DELIVERY_TIME',  # ReceivedTime
            'PR_CLIENT_SUBMIT_TIME',     # SentOn
            'PR_CREATION_TIME',
            'PR_MESSAGE_FLAGS',
        ]
        for tag_name in key_tags:
            if hasattr(mapitags, tag_name):
                value = getattr(mapitags, tag_name)
                print(f"  ✓ {tag_name} = 0x{value:08X}")
            else:
                print(f"  ❌ {tag_name} NOT FOUND")
                
    except ImportError as e:
        print(f"❌ mapitags not available: {e}")
    
    # =========================================================================
    # Step 4: Check for mapiutil
    # =========================================================================
    print_section("Step 4: Checking mapiutil")
    
    try:
        from win32com.mapi import mapiutil
        print(f"✓ mapiutil found")
        
        mapiutil_attrs = [a for a in dir(mapiutil) if not a.startswith('_')]
        print(f"\nAvailable in mapiutil ({len(mapiutil_attrs)} items):")
        for attr in sorted(mapiutil_attrs)[:20]:  # First 20
            print(f"  {attr}")
        if len(mapiutil_attrs) > 20:
            print(f"  ... and {len(mapiutil_attrs) - 20} more")
            
    except ImportError as e:
        print(f"❌ mapiutil not available: {e}")
    
    # =========================================================================
    # Step 5: Try basic MAPI initialization
    # =========================================================================
    print_section("Step 5: MAPI Initialization")
    
    try:
        from win32com.mapi import mapi
        
        # Check for MAPIInitialize
        if hasattr(mapi, 'MAPIInitialize'):
            print("✓ MAPIInitialize function exists")
            
            # Try to call it
            try:
                hr = mapi.MAPIInitialize(None)
                print(f"✓ MAPIInitialize() returned: {hr}")
                
                # Try to uninitialize
                if hasattr(mapi, 'MAPIUninitialize'):
                    mapi.MAPIUninitialize()
                    print("✓ MAPIUninitialize() called")
            except Exception as e:
                print(f"❌ MAPIInitialize failed: {e}")
        else:
            print("❌ MAPIInitialize not found in mapi module")
            
        # Check for MAPILogonEx
        if hasattr(mapi, 'MAPILogonEx'):
            print("✓ MAPILogonEx function exists")
        else:
            print("❌ MAPILogonEx not found")
            
    except Exception as e:
        print(f"❌ Error during MAPI init check: {e}")
    
    # =========================================================================
    # Step 6: Check for IConverterSession (EML import)
    # =========================================================================
    print_section("Step 6: IConverterSession (EML import capability)")
    
    try:
        from win32com.mapi import mapi
        
        if hasattr(mapi, 'CLSID_IConverterSession'):
            print(f"✓ CLSID_IConverterSession exists")
        else:
            print("❌ CLSID_IConverterSession not found")
            
        if hasattr(mapi, 'IID_IConverterSession'):
            print(f"✓ IID_IConverterSession exists")
        else:
            print("❌ IID_IConverterSession not found")
            
    except Exception as e:
        print(f"❌ Error checking IConverterSession: {e}")
    
    # =========================================================================
    # Step 7: Explore what COM interfaces might be available
    # =========================================================================
    print_section("Step 7: Checking for key MAPI interfaces")
    
    try:
        from win32com.mapi import mapi
        
        interfaces = [
            'IID_IMAPISession',
            'IID_IMsgStore', 
            'IID_IMAPIFolder',
            'IID_IMessage',
            'IID_IMAPIProp',
            'IID_IMAPITable',
        ]
        
        for iface in interfaces:
            if hasattr(mapi, iface):
                print(f"  ✓ {iface} defined")
            else:
                print(f"  ❌ {iface} NOT FOUND")
                
    except Exception as e:
        print(f"❌ Error checking interfaces: {e}")
    
    # =========================================================================
    # Summary
    # =========================================================================
    print_section("SUMMARY")
    print("""
Next steps depend on what we found:

1. If MAPIInitialize/MAPILogonEx exist and work:
   → We can try to use pywin32's MAPI bindings directly
   
2. If only basic stuff exists:
   → We may need to use comtypes to define missing interfaces
   
3. If IConverterSession exists:
   → There might be a way to import EML and set properties

Please paste this entire output back to me so I can analyze what's available!
""")

if __name__ == '__main__':
    main()
