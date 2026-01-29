"""
MAPI PST Creation - Milestone 3
===============================
Goal: Create a standalone PST file and write messages with custom dates

This will:
1. Create a new PST file (or use existing)
2. Add it to MAPI as a store
3. Create messages with historical dates
4. Verify dates are preserved
5. Optionally remove the PST from Outlook profile

Run with Outlook OPEN:
    python experiments/mapi_pst_create.py
"""

import sys
import os
from pathlib import Path

def print_section(title):
    print(f"\n{'='*60}")
    print(f" {title}")
    print('='*60)

def main():
    print("MAPI PST Creation - Milestone 3")
    
    if sys.platform != 'win32':
        print("\n❌ ERROR: This script must be run on Windows!")
        return
    
    # Output PST path - in Documents folder
    documents = Path(os.environ.get('USERPROFILE', '')) / 'Documents'
    pst_path = documents / 'MAPI_Test_Output.pst'
    
    print(f"Target PST: {pst_path}")
    
    # Import modules
    try:
        from win32com.mapi import mapi, mapitags
        import win32com.client
        import pythoncom
        import pywintypes
        from datetime import datetime
    except ImportError as e:
        print(f"❌ Failed to import modules: {e}")
        return
    
    session = None
    pst_store = None
    pst_root_eid = None
    
    try:
        # =====================================================================
        # Step 1: Initialize MAPI
        # =====================================================================
        print_section("Step 1: Initialize MAPI")
        
        pythoncom.CoInitialize()
        mapi.MAPIInitialize((mapi.MAPI_INIT_VERSION, mapi.MAPI_MULTITHREAD_NOTIFICATIONS))
        print("✓ MAPI initialized")
        
        # Connect to Outlook (needed for some operations)
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("✓ Connected to Outlook")
        
        # Get MAPI session
        session = mapi.MAPILogonEx(0, "", None, mapi.MAPI_EXTENDED | mapi.MAPI_USE_DEFAULT)
        print("✓ MAPI session established")
        
        # =====================================================================
        # Step 2: Create/Open PST file
        # =====================================================================
        print_section("Step 2: Create/Open PST File")
        
        # Remove existing file for clean test
        if pst_path.exists():
            print(f"Removing existing PST: {pst_path}")
            # First try to remove from Outlook if it's mounted
            try:
                for store in namespace.Stores:
                    if store.FilePath and Path(store.FilePath).resolve() == pst_path.resolve():
                        print(f"  Removing from Outlook profile first...")
                        namespace.RemoveStore(store.GetRootFolder())
                        break
            except:
                pass
            
            try:
                os.remove(pst_path)
                print("✓ Removed existing PST")
            except Exception as e:
                print(f"⚠ Could not remove: {e}")
        
        # Add PST to Outlook profile (creates if doesn't exist)
        print(f"\nAdding PST to profile: {pst_path}")
        namespace.AddStore(str(pst_path))
        print("✓ PST added to profile")
        
        # Give it a moment to initialize
        import time
        time.sleep(1)
        
        # Find the PST store we just added
        print("\nFinding new PST store...")
        pst_store = None
        for store in namespace.Stores:
            if store.FilePath:
                store_path = Path(store.FilePath).resolve()
                if store_path == pst_path.resolve():
                    pst_store = store
                    print(f"✓ Found PST store: {store.DisplayName}")
                    break
        
        if not pst_store:
            print("❌ Could not find the PST store we just added!")
            return
        
        # =====================================================================
        # Step 3: Open PST via MAPI
        # =====================================================================
        print_section("Step 3: Open PST via Extended MAPI")
        
        # Get the PST's entry ID via stores table
        stores_table = session.GetMsgStoresTable(0)
        
        PR_ENTRYID = mapitags.PR_ENTRYID
        PR_DISPLAY_NAME = mapitags.PR_DISPLAY_NAME_A
        PR_MDB_PROVIDER = 0x34140102  # Provider UID
        
        stores_table.SetColumns((PR_ENTRYID, PR_DISPLAY_NAME), 0)
        rows = stores_table.QueryRows(50, 0)
        
        pst_eid = None
        for row in rows:
            eid = row[0][1]
            name = row[1][1] if row[1][1] else b"(unknown)"
            if isinstance(name, bytes):
                name = name.decode('utf-8', errors='replace')
            
            # Match by name (the PST we added)
            if 'MAPI_Test' in name or pst_path.stem in name:
                pst_eid = eid
                print(f"✓ Found PST in stores table: {name}")
                break
        
        if not pst_eid:
            # Try by filepath match through Outlook
            pst_eid_hex = pst_store.StoreID
            print(f"Using StoreID from Outlook: {len(pst_eid_hex)} chars")
            # Convert hex string to bytes
            pst_eid = bytes.fromhex(pst_eid_hex)
        
        # Open the store via MAPI
        mapi_store = session.OpenMsgStore(0, pst_eid, None, mapi.MDB_WRITE | mapi.MAPI_BEST_ACCESS)
        print(f"✓ Opened PST via MAPI: {type(mapi_store)}")
        
        # =====================================================================
        # Step 4: Get/Create root folder
        # =====================================================================
        print_section("Step 4: Access Root Folder")
        
        # Get the IPM subtree (where mail folders live)
        PR_IPM_SUBTREE_ENTRYID = 0x35E00102
        
        props = mapi_store.GetProps([PR_IPM_SUBTREE_ENTRYID], 0)
        if props and props[0][0] == PR_IPM_SUBTREE_ENTRYID:
            ipm_subtree_eid = props[0][1]
            print(f"✓ Got IPM Subtree entry ID ({len(ipm_subtree_eid)} bytes)")
        else:
            print("⚠ Could not get IPM Subtree, using root")
            # Fallback: get root folder
            PR_ENTRYID = mapitags.PR_ENTRYID
            props = mapi_store.GetProps([PR_ENTRYID], 0)
            ipm_subtree_eid = props[0][1]
        
        # Open the IPM subtree folder
        root_folder = mapi_store.OpenEntry(ipm_subtree_eid, None, mapi.MAPI_MODIFY | mapi.MAPI_BEST_ACCESS)
        print(f"✓ Opened root folder: {type(root_folder)}")
        
        # =====================================================================
        # Step 5: Create "Imported Emails" subfolder
        # =====================================================================
        print_section("Step 5: Create Subfolder")
        
        # Create a subfolder for our imports
        folder_name = "Imported Emails"
        
        try:
            # CreateFolder(folderType, folderName, comment, IID, flags)
            # FOLDER_GENERIC = 1
            import_folder = root_folder.CreateFolder(1, folder_name, "Emails imported with custom dates", None, 0)
            print(f"✓ Created folder: {folder_name}")
        except Exception as e:
            if "already exists" in str(e).lower() or "MAPI_E_COLLISION" in str(e):
                print(f"Folder exists, opening it...")
                # Find and open existing folder
                contents = root_folder.GetHierarchyTable(0)
                contents.SetColumns([mapitags.PR_ENTRYID, mapitags.PR_DISPLAY_NAME_A], 0)
                rows = contents.QueryRows(100, 0)
                
                for row in rows:
                    fname = row[1][1]
                    if isinstance(fname, bytes):
                        fname = fname.decode('utf-8', errors='replace')
                    if fname == folder_name:
                        feid = row[0][1]
                        import_folder = mapi_store.OpenEntry(feid, None, mapi.MAPI_MODIFY | mapi.MAPI_BEST_ACCESS)
                        print(f"✓ Opened existing folder: {folder_name}")
                        break
            else:
                raise
        
        # =====================================================================
        # Step 6: Create test messages with different dates
        # =====================================================================
        print_section("Step 6: Create Test Messages")
        
        # Property tags we'll use
        PR_SUBJECT = mapitags.PR_SUBJECT_A
        PR_BODY = mapitags.PR_BODY_A
        PR_MESSAGE_CLASS = mapitags.PR_MESSAGE_CLASS_A
        PR_MESSAGE_FLAGS = mapitags.PR_MESSAGE_FLAGS
        PR_MESSAGE_DELIVERY_TIME = mapitags.PR_MESSAGE_DELIVERY_TIME
        PR_CLIENT_SUBMIT_TIME = mapitags.PR_CLIENT_SUBMIT_TIME
        PR_SENDER_NAME = mapitags.PR_SENDER_NAME_A
        PR_SENDER_EMAIL_ADDRESS = mapitags.PR_SENDER_EMAIL_ADDRESS_A
        PR_SENT_REPRESENTING_NAME = 0x0042001E  # PR_SENT_REPRESENTING_NAME_A
        PR_SENT_REPRESENTING_EMAIL = 0x0065001E  # PR_SENT_REPRESENTING_EMAIL_ADDRESS_A
        
        # MSGFLAG_READ = 1
        MSGFLAG_READ = 0x0001
        
        # Test messages with different dates
        test_messages = [
            {
                "subject": "Email from 2015",
                "body": "This email is dated January 15, 2015",
                "date": datetime(2015, 1, 15, 9, 30, 0),
                "sender_name": "Alice Smith",
                "sender_email": "alice@example.com",
            },
            {
                "subject": "Email from 2018",
                "body": "This email is dated July 4, 2018",
                "date": datetime(2018, 7, 4, 14, 0, 0),
                "sender_name": "Bob Johnson",
                "sender_email": "bob@example.com",
            },
            {
                "subject": "Email from 2020",
                "body": "This email is dated March 15, 2020",
                "date": datetime(2020, 3, 15, 11, 45, 0),
                "sender_name": "Carol Williams",
                "sender_email": "carol@example.com",
            },
            {
                "subject": "Email from last year",
                "body": "This email is dated June 1, 2025",
                "date": datetime(2025, 6, 1, 8, 0, 0),
                "sender_name": "David Brown",
                "sender_email": "david@example.com",
            },
        ]
        
        created_count = 0
        for msg_data in test_messages:
            try:
                # Create message
                msg = import_folder.CreateMessage(None, 0)
                
                # Convert date to PyTime
                pytime = pywintypes.Time(msg_data["date"])
                
                # Set properties
                props = [
                    (PR_MESSAGE_CLASS, "IPM.Note"),
                    (PR_SUBJECT, msg_data["subject"]),
                    (PR_BODY, msg_data["body"]),
                    (PR_MESSAGE_FLAGS, MSGFLAG_READ),
                    (PR_MESSAGE_DELIVERY_TIME, pytime),
                    (PR_CLIENT_SUBMIT_TIME, pytime),
                    (PR_SENDER_NAME, msg_data["sender_name"]),
                    (PR_SENDER_EMAIL_ADDRESS, msg_data["sender_email"]),
                    (PR_SENT_REPRESENTING_NAME, msg_data["sender_name"]),
                    (PR_SENT_REPRESENTING_EMAIL, msg_data["sender_email"]),
                ]
                
                msg.SetProps(props)
                msg.SaveChanges(0)
                
                print(f"✓ Created: '{msg_data['subject']}' - Date: {msg_data['date']}")
                created_count += 1
                
            except Exception as e:
                print(f"❌ Failed to create '{msg_data['subject']}': {e}")
        
        print(f"\n✓ Created {created_count}/{len(test_messages)} messages")
        
        # =====================================================================
        # Summary
        # =====================================================================
        print_section("SUCCESS!")
        print(f"""
PST file created: {pst_path}

The PST should now contain a folder called "Imported Emails" with 4 test 
messages, each with different historical dates:
  - January 15, 2015
  - July 4, 2018  
  - March 15, 2020
  - June 1, 2025

CHECK IN OUTLOOK:
1. Look for "Personal Folders" or similar in your folder list
2. Open the "Imported Emails" folder
3. Verify the dates shown are the HISTORICAL dates, not today!

The PST is still mounted in Outlook. To safely remove it:
  - Right-click on the PST in Outlook
  - Select "Close" or "Remove"

Or run this script again - it will clean up and recreate.
""")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        # =====================================================================
        # Cleanup MAPI (but leave PST mounted for inspection)
        # =====================================================================
        print_section("Cleanup")
        try:
            mapi.MAPIUninitialize()
            pythoncom.CoUninitialize()
            print("✓ MAPI cleanup completed")
            print(f"\nPST remains mounted at: {pst_path}")
        except:
            pass

if __name__ == '__main__':
    main()
