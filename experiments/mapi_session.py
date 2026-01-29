"""
MAPI Session & PST Access - Milestone 2
=======================================
Goal: Log into MAPI and access/create a PST store

Run this on Windows with Outlook installed:
    python experiments/mapi_session.py

Prerequisites:
- Outlook must be installed
- You should have at least one mail profile configured
"""

import sys
import os

def print_section(title):
    print(f"\n{'='*60}")
    print(f" {title}")
    print('='*60)

def main():
    print("MAPI Session Exploration - Milestone 2")
    
    if sys.platform != 'win32':
        print("\n❌ ERROR: This script must be run on Windows!")
        return
    
    # Import MAPI modules
    try:
        from win32com.mapi import mapi, mapitags, mapiutil
        import pythoncom
    except ImportError as e:
        print(f"❌ Failed to import MAPI modules: {e}")
        return
    
    # =========================================================================
    # Step 1: Initialize MAPI
    # =========================================================================
    print_section("Step 1: MAPI Initialization")
    
    try:
        # Initialize with multithread notifications
        mapi.MAPIInitialize((mapi.MAPI_INIT_VERSION, mapi.MAPI_MULTITHREAD_NOTIFICATIONS))
        print("✓ MAPIInitialize() succeeded")
    except Exception as e:
        print(f"❌ MAPIInitialize failed: {e}")
        return
    
    try:
        # =====================================================================
        # Step 2: Log into MAPI Session
        # =====================================================================
        print_section("Step 2: MAPI Logon")
        
        session = None
        try:
            # Try to logon - MAPI_USE_DEFAULT uses default profile
            # Flags: MAPI_EXTENDED | MAPI_USE_DEFAULT | MAPI_NO_MAIL
            flags = mapi.MAPI_EXTENDED | mapi.MAPI_USE_DEFAULT
            
            print(f"Attempting MAPILogonEx with flags: {flags}")
            session = mapi.MAPILogonEx(0, None, None, flags)
            print(f"✓ MAPILogonEx succeeded!")
            print(f"  Session object type: {type(session)}")
            print(f"  Session: {session}")
            
        except Exception as e:
            print(f"❌ MAPILogonEx failed: {e}")
            print("\nTrying alternative: Connect via Outlook...")
            
            # Alternative: Use Outlook's session
            try:
                import win32com.client
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                print(f"✓ Connected to Outlook")
                print(f"  Namespace: {namespace}")
                
                # Try to get MAPIOBJECT
                if hasattr(namespace, 'MAPIOBJECT'):
                    mapi_obj = namespace.MAPIOBJECT
                    print(f"✓ Got MAPIOBJECT: {type(mapi_obj)}")
                else:
                    print("❌ MAPIOBJECT not accessible")
                    
            except Exception as e2:
                print(f"❌ Outlook connection also failed: {e2}")
                return
        
        # =====================================================================
        # Step 3: List available message stores
        # =====================================================================
        print_section("Step 3: List Message Stores")
        
        if session:
            try:
                # GetMsgStoresTable returns a table of all stores
                stores_table = session.GetMsgStoresTable(0)
                print(f"✓ Got stores table: {type(stores_table)}")
                
                # Try to query rows
                # We need PR_DISPLAY_NAME and PR_ENTRYID
                PR_DISPLAY_NAME = mapitags.PR_DISPLAY_NAME_A  # 0x3001001E
                PR_ENTRYID = mapitags.PR_ENTRYID  # 0x0FFF0102
                PR_DEFAULT_STORE = mapitags.PR_DEFAULT_STORE  # 0x3400000B
                
                # Set columns
                stores_table.SetColumns((PR_ENTRYID, PR_DISPLAY_NAME, PR_DEFAULT_STORE), 0)
                
                # Query all rows
                rows = stores_table.QueryRows(100, 0)
                print(f"✓ Found {len(rows)} message stores:")
                
                for i, row in enumerate(rows):
                    entry_id = row[0][1]  # PR_ENTRYID value
                    display_name = row[1][1] if row[1][1] else "(unknown)"
                    is_default = row[2][1] if len(row) > 2 and row[2][1] else False
                    
                    default_marker = " [DEFAULT]" if is_default else ""
                    print(f"  {i+1}. {display_name}{default_marker}")
                    
            except Exception as e:
                print(f"❌ Error listing stores: {e}")
                import traceback
                traceback.print_exc()
        
        # =====================================================================
        # Step 4: Try to open a message store
        # =====================================================================
        print_section("Step 4: Open a Message Store")
        
        if session:
            try:
                # Re-query to get entry IDs
                stores_table = session.GetMsgStoresTable(0)
                PR_ENTRYID = mapitags.PR_ENTRYID
                PR_DISPLAY_NAME = mapitags.PR_DISPLAY_NAME_A
                stores_table.SetColumns((PR_ENTRYID, PR_DISPLAY_NAME), 0)
                rows = stores_table.QueryRows(100, 0)
                
                if rows:
                    # Open the first store
                    entry_id = rows[0][0][1]
                    store_name = rows[0][1][1] if rows[0][1][1] else "(unknown)"
                    
                    print(f"Opening store: {store_name}")
                    
                    # OpenMsgStore(ulUIParam, entryID, IID, flags)
                    store = session.OpenMsgStore(
                        0,  # UI param
                        entry_id,
                        None,  # IID (None = IMsgStore)
                        mapi.MDB_WRITE | mapi.MAPI_BEST_ACCESS
                    )
                    print(f"✓ Opened store: {type(store)}")
                    
                    # List methods available on store
                    store_methods = [m for m in dir(store) if not m.startswith('_')]
                    print(f"\nStore methods ({len(store_methods)}):")
                    for m in store_methods[:15]:
                        print(f"    {m}")
                    if len(store_methods) > 15:
                        print(f"    ... and {len(store_methods) - 15} more")
                    
            except Exception as e:
                print(f"❌ Error opening store: {e}")
                import traceback
                traceback.print_exc()
        
        # =====================================================================
        # Step 5: Try to get Inbox folder
        # =====================================================================
        print_section("Step 5: Access Inbox Folder")
        
        if session:
            try:
                # Re-open store
                stores_table = session.GetMsgStoresTable(0)
                stores_table.SetColumns((mapitags.PR_ENTRYID,), 0)
                rows = stores_table.QueryRows(1, 0)
                
                if rows:
                    entry_id = rows[0][0][1]
                    store = session.OpenMsgStore(0, entry_id, None, mapi.MDB_WRITE | mapi.MAPI_BEST_ACCESS)
                    
                    # Get Inbox entry ID
                    # GetReceiveFolder returns (entry_id, message_class)
                    inbox_eid, _ = store.GetReceiveFolder("IPM.Note", 0)
                    print(f"✓ Got Inbox entry ID: {len(inbox_eid)} bytes")
                    
                    # Open Inbox folder
                    inbox = store.OpenEntry(
                        inbox_eid,
                        None,  # IID
                        mapi.MAPI_BEST_ACCESS | mapi.MAPI_MODIFY
                    )
                    print(f"✓ Opened Inbox: {type(inbox)}")
                    
                    # List folder methods
                    folder_methods = [m for m in dir(inbox) if not m.startswith('_')]
                    print(f"\nFolder methods ({len(folder_methods)}):")
                    for m in folder_methods:
                        print(f"    {m}")
                    
                    # Check if CreateMessage exists!
                    if 'CreateMessage' in folder_methods:
                        print("\n🎉 CreateMessage EXISTS! We can create messages!")
                    else:
                        print("\n⚠️ CreateMessage not found in folder methods")
                        
            except Exception as e:
                print(f"❌ Error accessing Inbox: {e}")
                import traceback
                traceback.print_exc()
        
        # =====================================================================
        # Step 6: Explore IMessage creation (if possible)
        # =====================================================================
        print_section("Step 6: Test Message Creation")
        
        if session:
            try:
                # Re-open everything
                stores_table = session.GetMsgStoresTable(0)
                stores_table.SetColumns((mapitags.PR_ENTRYID,), 0)
                rows = stores_table.QueryRows(1, 0)
                
                if rows:
                    entry_id = rows[0][0][1]
                    store = session.OpenMsgStore(0, entry_id, None, mapi.MDB_WRITE | mapi.MAPI_BEST_ACCESS)
                    inbox_eid, _ = store.GetReceiveFolder("IPM.Note", 0)
                    inbox = store.OpenEntry(inbox_eid, None, mapi.MAPI_BEST_ACCESS | mapi.MAPI_MODIFY)
                    
                    # Try to create a message
                    print("Attempting to create a message...")
                    msg = inbox.CreateMessage(None, 0)
                    print(f"✓ Created message: {type(msg)}")
                    
                    # List message methods
                    msg_methods = [m for m in dir(msg) if not m.startswith('_')]
                    print(f"\nMessage methods ({len(msg_methods)}):")
                    for m in msg_methods:
                        print(f"    {m}")
                    
                    # KEY: Check for SetProps!
                    if 'SetProps' in msg_methods:
                        print("\n🎉🎉🎉 SetProps EXISTS! We can set properties including DATES!")
                    else:
                        print("\n⚠️ SetProps not found - will need alternative approach")
                    
                    # Don't save - just exploring
                    print("\n(Message not saved - this was just a test)")
                    
            except Exception as e:
                print(f"❌ Error creating message: {e}")
                import traceback
                traceback.print_exc()
    
    finally:
        # =====================================================================
        # Cleanup
        # =====================================================================
        print_section("Cleanup")
        try:
            mapi.MAPIUninitialize()
            print("✓ MAPIUninitialize() called")
        except:
            pass
    
    # =========================================================================
    # Summary
    # =========================================================================
    print_section("SUMMARY")
    print("""
What we discovered:

1. If MAPILogonEx worked and we could list stores:
   → We have full MAPI access!
   
2. If CreateMessage exists on folder:
   → We can create new messages
   
3. If SetProps exists on message:
   → We can set PR_MESSAGE_DELIVERY_TIME and PR_CLIENT_SUBMIT_TIME!
   
Next milestone: Actually set date properties and save a message.

Please paste this output back to me!
""")

if __name__ == '__main__':
    main()
