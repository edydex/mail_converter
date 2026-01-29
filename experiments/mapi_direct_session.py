"""
MAPI Direct Session - Milestone 2c
==================================
Goal: Use MAPI session directly without QueryInterface issues

This approach:
1. Gets IMAPISession from Outlook's Namespace.MAPIOBJECT
2. Uses OpenMsgStore to open stores directly
3. Avoids the problematic QueryInterface on folder MAPIObjects
"""

import sys
import os

def print_section(title):
    print(f"\n{'='*60}")
    print(f" {title}")
    print('='*60)

def main():
    print("MAPI Direct Session - Milestone 2c")
    
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
    # Step 1: Initialize and get session
    # =========================================================================
    print_section("Step 1: Initialize MAPI and get session")
    
    session = None
    try:
        pythoncom.CoInitialize()
        
        # Initialize MAPI
        mapi.MAPIInitialize((mapi.MAPI_INIT_VERSION, mapi.MAPI_MULTITHREAD_NOTIFICATIONS))
        print("✓ MAPI initialized")
        
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("✓ Connected to Outlook")
        
        # Get IMAPISession from Outlook
        mapi_obj = namespace.MAPIOBJECT
        print(f"✓ Got MAPIOBJECT: {type(mapi_obj)}")
        
        # QueryInterface to IMAPISession
        session = mapi_obj.QueryInterface(mapi.IID_IMAPISession)
        print(f"✓ Got IMAPISession: {type(session)}")
        
        # List session methods
        session_methods = [m for m in dir(session) if not m.startswith('_')]
        print(f"\nIMAPISession methods ({len(session_methods)}):")
        for m in session_methods:
            print(f"    {m}")
            
    except Exception as e:
        print(f"❌ Failed to get session: {e}")
        import traceback
        traceback.print_exc()
        
        # Try alternative: direct logon with profile
        print("\n--- Trying alternative: MAPILogonEx with profile ---")
        try:
            # Get profile name from Outlook
            profile_name = namespace.CurrentProfileName if hasattr(namespace, 'CurrentProfileName') else None
            print(f"Current profile: {profile_name}")
            
            # Try different flag combinations
            flag_combos = [
                ("MAPI_EXTENDED | MAPI_USE_DEFAULT", mapi.MAPI_EXTENDED | mapi.MAPI_USE_DEFAULT),
                ("MAPI_EXTENDED | MAPI_NEW_SESSION", mapi.MAPI_EXTENDED | mapi.MAPI_NEW_SESSION),
                ("MAPI_EXTENDED", mapi.MAPI_EXTENDED),
            ]
            
            for flag_name, flags in flag_combos:
                try:
                    print(f"\nTrying MAPILogonEx with {flag_name}...")
                    session = mapi.MAPILogonEx(0, "", None, flags)
                    print(f"✓ MAPILogonEx succeeded with {flag_name}!")
                    break
                except Exception as e2:
                    print(f"  Failed: {e2}")
                    
        except Exception as e3:
            print(f"❌ Alternative also failed: {e3}")
    
    if not session:
        print("\n❌ Could not get MAPI session. Trying one more approach...")
        
        # =====================================================================
        # Alternative: Create message via Outlook, get its MAPI object
        # =====================================================================
        print_section("Alternative: Create via Outlook, modify via MAPI")
        
        try:
            # Create a mail item via Outlook
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.Subject = "Test - Will set date via MAPI"
            mail.Body = "This is a test message"
            mail.Save()  # Save to Drafts first
            print("✓ Created and saved draft via Outlook")
            
            # Get its MAPI object
            mail_mapi = mail.MAPIOBJECT
            print(f"✓ Got mail MAPIOBJECT: {type(mail_mapi)}")
            
            # Try QueryInterface to IMessage
            try:
                msg = mail_mapi.QueryInterface(mapi.IID_IMessage)
                print(f"✓ QueryInterface to IMessage succeeded: {type(msg)}")
                
                # List methods
                msg_methods = [m for m in dir(msg) if not m.startswith('_')]
                print(f"\nIMessage methods ({len(msg_methods)}):")
                for m in msg_methods:
                    print(f"    {m}")
                    
                if 'SetProps' in msg_methods:
                    print("\n🎉 SetProps EXISTS!")
                    
            except Exception as qi_err:
                print(f"❌ QueryInterface to IMessage failed: {qi_err}")
                
            # Clean up - delete the draft
            mail.Delete()
            print("✓ Deleted test draft")
            
        except Exception as alt_err:
            print(f"❌ Alternative approach failed: {alt_err}")
            import traceback
            traceback.print_exc()
        
        return
    
    # =========================================================================
    # Step 2: Get message stores table
    # =========================================================================
    print_section("Step 2: Get Message Stores")
    
    try:
        stores_table = session.GetMsgStoresTable(0)
        print(f"✓ Got stores table: {type(stores_table)}")
        
        # Set columns
        PR_ENTRYID = mapitags.PR_ENTRYID
        PR_DISPLAY_NAME = mapitags.PR_DISPLAY_NAME_A
        PR_DEFAULT_STORE = mapitags.PR_DEFAULT_STORE
        
        stores_table.SetColumns((PR_ENTRYID, PR_DISPLAY_NAME, PR_DEFAULT_STORE), 0)
        
        # Query rows
        rows = stores_table.QueryRows(50, 0)
        print(f"✓ Found {len(rows)} stores:")
        
        default_store_eid = None
        for row in rows:
            eid = row[0][1]
            name = row[1][1] if row[1][0] == PR_DISPLAY_NAME else "(unknown)"
            is_default = row[2][1] if len(row) > 2 and row[2][0] == PR_DEFAULT_STORE else False
            
            marker = " [DEFAULT]" if is_default else ""
            print(f"    {name}{marker}")
            
            if is_default:
                default_store_eid = eid
                
    except Exception as e:
        print(f"❌ Error getting stores: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # =========================================================================
    # Step 3: Open default store
    # =========================================================================
    print_section("Step 3: Open Default Store")
    
    store = None
    try:
        if default_store_eid:
            store = session.OpenMsgStore(0, default_store_eid, None, mapi.MDB_WRITE | mapi.MAPI_BEST_ACCESS)
            print(f"✓ Opened default store: {type(store)}")
        else:
            # Just open first store
            stores_table = session.GetMsgStoresTable(0)
            stores_table.SetColumns((PR_ENTRYID,), 0)
            rows = stores_table.QueryRows(1, 0)
            if rows:
                store = session.OpenMsgStore(0, rows[0][0][1], None, mapi.MDB_WRITE | mapi.MAPI_BEST_ACCESS)
                print(f"✓ Opened first store: {type(store)}")
                
    except Exception as e:
        print(f"❌ Error opening store: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # =========================================================================
    # Step 4: Get Inbox folder
    # =========================================================================
    print_section("Step 4: Get Inbox Folder")
    
    inbox = None
    try:
        # GetReceiveFolder returns the inbox entry ID
        inbox_eid, _ = store.GetReceiveFolder("IPM.Note", 0)
        print(f"✓ Got Inbox entry ID ({len(inbox_eid)} bytes)")
        
        # Open the folder
        inbox = store.OpenEntry(inbox_eid, None, mapi.MAPI_MODIFY | mapi.MAPI_BEST_ACCESS)
        print(f"✓ Opened Inbox: {type(inbox)}")
        
        # List methods
        inbox_methods = [m for m in dir(inbox) if not m.startswith('_')]
        print(f"\nInbox methods ({len(inbox_methods)}):")
        for m in inbox_methods:
            print(f"    {m}")
            
        if 'CreateMessage' in inbox_methods:
            print("\n🎉 CreateMessage EXISTS!")
            
    except Exception as e:
        print(f"❌ Error getting Inbox: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # =========================================================================
    # Step 5: Create message and set dates
    # =========================================================================
    print_section("Step 5: Create Message with Custom Dates")
    
    try:
        from datetime import datetime
        import pywintypes
        
        # Create message
        msg = inbox.CreateMessage(None, 0)
        print(f"✓ Created message: {type(msg)}")
        
        # List methods
        msg_methods = [m for m in dir(msg) if not m.startswith('_')]
        print(f"Message methods: {msg_methods}")
        
        if 'SetProps' not in msg_methods:
            print("❌ SetProps not available!")
            return
            
        print("✓ SetProps is available!")
        
        # Prepare properties
        PR_SUBJECT = mapitags.PR_SUBJECT_A
        PR_BODY = mapitags.PR_BODY_A
        PR_MESSAGE_CLASS = mapitags.PR_MESSAGE_CLASS_A
        PR_MESSAGE_FLAGS = mapitags.PR_MESSAGE_FLAGS
        PR_MESSAGE_DELIVERY_TIME = mapitags.PR_MESSAGE_DELIVERY_TIME
        PR_CLIENT_SUBMIT_TIME = mapitags.PR_CLIENT_SUBMIT_TIME
        
        # Test date - June 15, 2020
        test_date = datetime(2020, 6, 15, 14, 30, 0)
        pytime = pywintypes.Time(test_date)
        
        # MSGFLAG_READ = 1 (mark as read, not unsent)
        MSGFLAG_READ = 0x0001
        
        print(f"\nSetting properties...")
        print(f"  Subject: 'MAPI Date Test - Should show June 15, 2020'")
        print(f"  Date: {test_date}")
        
        props = [
            (PR_MESSAGE_CLASS, "IPM.Note"),
            (PR_SUBJECT, "MAPI Date Test - Should show June 15, 2020"),
            (PR_BODY, "This message was created via Extended MAPI.\n\nThe date should show June 15, 2020, NOT today's date."),
            (PR_MESSAGE_FLAGS, MSGFLAG_READ),
            (PR_MESSAGE_DELIVERY_TIME, pytime),
            (PR_CLIENT_SUBMIT_TIME, pytime),
        ]
        
        result = msg.SetProps(props)
        print(f"✓ SetProps result: {result}")
        
        # Save BEFORE reading back (important for MAPI)
        print("\nSaving message...")
        msg.SaveChanges(mapi.KEEP_OPEN_READWRITE)
        print("✓ Message saved!")
        
        # Read back dates to verify
        read_props = msg.GetProps([PR_MESSAGE_DELIVERY_TIME, PR_CLIENT_SUBMIT_TIME], 0)
        print(f"\nDates after save:")
        for prop in read_props:
            tag, value = prop
            if tag == PR_MESSAGE_DELIVERY_TIME:
                print(f"  PR_MESSAGE_DELIVERY_TIME: {value}")
            elif tag == PR_CLIENT_SUBMIT_TIME:
                print(f"  PR_CLIENT_SUBMIT_TIME: {value}")
        
        print("\n" + "="*60)
        print(" 🎉 CHECK YOUR OUTLOOK INBOX! 🎉")
        print(" Look for: 'MAPI Date Test - Should show June 15, 2020'")
        print(" The received date should show June 15, 2020, NOT today!")
        print("="*60)
        
    except Exception as e:
        print(f"❌ Error creating message: {e}")
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
