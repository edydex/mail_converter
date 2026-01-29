"""
MAPI Session via Outlook - Milestone 2b
=======================================
Goal: Use Outlook's MAPI session to access stores and create messages

Run this on Windows with Outlook OPEN:
    python experiments/mapi_session_outlook.py

This version uses Outlook's existing session instead of MAPILogonEx.
"""

import sys
import os
import time

def print_section(title):
    print(f"\n{'='*60}")
    print(f" {title}")
    print('='*60)

def main():
    print("MAPI Session via Outlook - Milestone 2b")
    
    if sys.platform != 'win32':
        print("\n❌ ERROR: This script must be run on Windows!")
        return
    
    # Import modules
    try:
        from win32com.mapi import mapi, mapitags
        import win32com.client
        import pythoncom
    except ImportError as e:
        print(f"❌ Failed to import modules: {e}")
        return
    
    # =========================================================================
    # Step 1: Connect to Outlook and get MAPI session
    # =========================================================================
    print_section("Step 1: Connect to Outlook")
    
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print(f"✓ Connected to Outlook")
        
        # Get the MAPI object - this is the IMAPISession
        mapi_session = namespace.MAPIOBJECT
        print(f"✓ Got MAPIOBJECT: {type(mapi_session)}")
        
    except Exception as e:
        print(f"❌ Failed to connect to Outlook: {e}")
        print("\nMake sure Outlook is running!")
        return
    
    # =========================================================================
    # Step 2: List stores via Outlook (easier approach)
    # =========================================================================
    print_section("Step 2: List Message Stores (via Outlook)")
    
    try:
        stores = namespace.Stores
        print(f"✓ Found {stores.Count} message stores:")
        
        for i in range(1, stores.Count + 1):
            store = stores.Item(i)
            print(f"  {i}. {store.DisplayName}")
            print(f"     FilePath: {store.FilePath if store.FilePath else '(Exchange/Cloud)'}")
            
    except Exception as e:
        print(f"❌ Error listing stores: {e}")
    
    # =========================================================================
    # Step 3: Try to access Inbox via MAPI directly
    # =========================================================================
    print_section("Step 3: Access Inbox via Extended MAPI")
    
    inbox_folder = None
    try:
        # Initialize MAPI 
        mapi.MAPIInitialize((mapi.MAPI_INIT_VERSION, mapi.MAPI_MULTITHREAD_NOTIFICATIONS))
        print("✓ MAPI initialized")
        
        # Get Inbox from Outlook
        inbox_outlook = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        print(f"✓ Got Inbox via Outlook: {inbox_outlook.Name}")
        
        # Get the MAPI object for the folder
        inbox_mapi = inbox_outlook.MAPIOBJECT
        print(f"✓ Got Inbox MAPIOBJECT: {type(inbox_mapi)}")
        
        # This is a PyIUnknown - we need to QueryInterface to IMAPIFolder
        # Let's see what methods it has
        print(f"\nInbox MAPIOBJECT methods:")
        for m in dir(inbox_mapi):
            if not m.startswith('_'):
                print(f"    {m}")
        
    except Exception as e:
        print(f"❌ Error accessing Inbox MAPI: {e}")
        import traceback
        traceback.print_exc()
    
    # =========================================================================
    # Step 4: Try QueryInterface to get IMAPIFolder
    # =========================================================================
    print_section("Step 4: QueryInterface to IMAPIFolder")
    
    try:
        inbox_outlook = namespace.GetDefaultFolder(6)
        inbox_mapi = inbox_outlook.MAPIOBJECT
        
        # QueryInterface to IMAPIFolder
        IID_IMAPIFolder = mapi.IID_IMAPIFolder
        print(f"IID_IMAPIFolder: {IID_IMAPIFolder}")
        
        folder = inbox_mapi.QueryInterface(IID_IMAPIFolder)
        print(f"✓ QueryInterface succeeded: {type(folder)}")
        
        # List methods on the folder
        folder_methods = [m for m in dir(folder) if not m.startswith('_')]
        print(f"\nIMAPIFolder methods ({len(folder_methods)}):")
        for m in folder_methods:
            print(f"    {m}")
        
        # Check for CreateMessage
        if 'CreateMessage' in folder_methods:
            print("\n🎉 CreateMessage EXISTS!")
        
    except Exception as e:
        print(f"❌ QueryInterface failed: {e}")
        import traceback
        traceback.print_exc()
    
    # =========================================================================
    # Step 5: Create a message and explore its methods
    # =========================================================================
    print_section("Step 5: Create Message")
    
    try:
        inbox_outlook = namespace.GetDefaultFolder(6)
        inbox_mapi = inbox_outlook.MAPIOBJECT
        folder = inbox_mapi.QueryInterface(mapi.IID_IMAPIFolder)
        
        # Create a message
        print("Creating message...")
        msg = folder.CreateMessage(None, 0)
        print(f"✓ Created message: {type(msg)}")
        
        # List message methods
        msg_methods = [m for m in dir(msg) if not m.startswith('_')]
        print(f"\nIMessage methods ({len(msg_methods)}):")
        for m in msg_methods:
            print(f"    {m}")
        
        # KEY CHECKS
        print("\n*** KEY METHOD CHECKS ***")
        if 'SetProps' in msg_methods:
            print("✓ SetProps - CAN SET PROPERTIES!")
        else:
            print("❌ SetProps not found")
            
        if 'SaveChanges' in msg_methods:
            print("✓ SaveChanges - CAN SAVE!")
        else:
            print("❌ SaveChanges not found")
            
        if 'GetProps' in msg_methods:
            print("✓ GetProps - CAN READ PROPERTIES!")
        else:
            print("❌ GetProps not found")
        
        # Don't save - just exploring
        print("\n(Message not saved - this was just a test)")
        
    except Exception as e:
        print(f"❌ Error creating message: {e}")
        import traceback
        traceback.print_exc()
    
    # =========================================================================
    # Step 6: Test SetProps with a simple property
    # =========================================================================
    print_section("Step 6: Test SetProps")
    
    try:
        inbox_outlook = namespace.GetDefaultFolder(6)
        inbox_mapi = inbox_outlook.MAPIOBJECT
        folder = inbox_mapi.QueryInterface(mapi.IID_IMAPIFolder)
        
        # Create a fresh message
        msg = folder.CreateMessage(None, 0)
        print("✓ Created new message")
        
        # Try to set subject using SetProps
        PR_SUBJECT = mapitags.PR_SUBJECT_A  # 0x0037001E (string)
        
        print(f"\nTrying SetProps with PR_SUBJECT (0x{PR_SUBJECT:08X})...")
        
        # SetProps takes a sequence of (tag, value) tuples
        props = [(PR_SUBJECT, "Test Subject from MAPI")]
        
        result = msg.SetProps(props)
        print(f"✓ SetProps returned: {result}")
        
        # Read it back
        read_props = msg.GetProps([PR_SUBJECT], 0)
        print(f"✓ GetProps returned: {read_props}")
        
        if read_props and read_props[0][1]:
            print(f"✓ Subject value: {read_props[0][1]}")
        
        print("\n✓ SetProps WORKS! We can set properties!")
        
    except Exception as e:
        print(f"❌ SetProps test failed: {e}")
        import traceback
        traceback.print_exc()
    
    # =========================================================================
    # Step 7: Test setting DATE property (THE KEY TEST!)
    # =========================================================================
    print_section("Step 7: TEST DATE PROPERTY (THE KEY TEST!)")
    
    try:
        from datetime import datetime
        import pywintypes
        
        inbox_outlook = namespace.GetDefaultFolder(6)
        inbox_mapi = inbox_outlook.MAPIOBJECT
        folder = inbox_mapi.QueryInterface(mapi.IID_IMAPIFolder)
        
        # Create a fresh message
        msg = folder.CreateMessage(None, 0)
        print("✓ Created new message")
        
        # The key property tags
        PR_MESSAGE_DELIVERY_TIME = mapitags.PR_MESSAGE_DELIVERY_TIME  # 0x0E060040
        PR_CLIENT_SUBMIT_TIME = mapitags.PR_CLIENT_SUBMIT_TIME        # 0x00390040
        PR_MESSAGE_FLAGS = mapitags.PR_MESSAGE_FLAGS                  # 0x0E070003
        
        # MSGFLAG_READ = 1, MSGFLAG_UNSENT = 8
        # To mark as "sent/received", we clear MSGFLAG_UNSENT
        MSGFLAG_READ = 0x0001
        
        # Create a test date - let's use a specific past date
        test_date = datetime(2020, 6, 15, 14, 30, 0)  # June 15, 2020, 2:30 PM
        
        # Convert to PyTime (what MAPI expects for PT_SYSTIME)
        pytime = pywintypes.Time(test_date)
        print(f"Test date: {test_date}")
        print(f"PyTime: {pytime}")
        
        # Set properties BEFORE first save
        print(f"\nSetting PR_MESSAGE_DELIVERY_TIME (0x{PR_MESSAGE_DELIVERY_TIME:08X})...")
        print(f"Setting PR_CLIENT_SUBMIT_TIME (0x{PR_CLIENT_SUBMIT_TIME:08X})...")
        
        props = [
            (PR_MESSAGE_DELIVERY_TIME, pytime),
            (PR_CLIENT_SUBMIT_TIME, pytime),
            (PR_MESSAGE_FLAGS, MSGFLAG_READ),  # Mark as read, not unsent
            (mapitags.PR_SUBJECT_A, "MAPI Date Test - Should show June 15, 2020"),
        ]
        
        result = msg.SetProps(props)
        print(f"✓ SetProps returned: {result}")
        
        # Read back the dates
        read_props = msg.GetProps([PR_MESSAGE_DELIVERY_TIME, PR_CLIENT_SUBMIT_TIME], 0)
        print(f"\nReading back dates:")
        for tag, value in read_props:
            tag_name = "PR_MESSAGE_DELIVERY_TIME" if tag == PR_MESSAGE_DELIVERY_TIME else "PR_CLIENT_SUBMIT_TIME"
            print(f"  {tag_name}: {value}")
        
        # NOW THE BIG TEST: Save and check if dates persist
        print("\n*** SAVING MESSAGE ***")
        msg.SaveChanges(mapi.KEEP_OPEN_READWRITE)
        print("✓ SaveChanges completed!")
        
        # Read dates after save
        read_props_after = msg.GetProps([PR_MESSAGE_DELIVERY_TIME, PR_CLIENT_SUBMIT_TIME], 0)
        print(f"\nDates AFTER save:")
        for tag, value in read_props_after:
            tag_name = "PR_MESSAGE_DELIVERY_TIME" if tag == PR_MESSAGE_DELIVERY_TIME else "PR_CLIENT_SUBMIT_TIME"
            print(f"  {tag_name}: {value}")
        
        print("\n" + "="*60)
        print(" CHECK YOUR OUTLOOK INBOX!")
        print(" Look for: 'MAPI Date Test - Should show June 15, 2020'")
        print(" The received date should show June 15, 2020, NOT today!")
        print("="*60)
        
    except Exception as e:
        print(f"❌ Date property test failed: {e}")
        import traceback
        traceback.print_exc()
    
    # =========================================================================
    # Cleanup
    # =========================================================================
    print_section("Cleanup")
    try:
        mapi.MAPIUninitialize()
        pythoncom.CoUninitialize()
        print("✓ Cleanup completed")
    except:
        pass

if __name__ == '__main__':
    main()
