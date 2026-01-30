"""
Minimal test - copy EXACT pattern from working mapi_pst_create.py
to verify the approach works.
"""

import sys
import os
from pathlib import Path

def main():
    print("Minimal MAPI Test - copying exact working pattern")
    
    if sys.platform != 'win32':
        print("❌ Windows only!")
        return
    
    # Output PST
    documents = Path(os.environ.get('USERPROFILE', '')) / 'Documents'
    pst_path = documents / 'MAPI_Minimal_Test.pst'
    
    print(f"Target PST: {pst_path}")
    
    # Import modules - EXACTLY like working script
    from win32com.mapi import mapi, mapitags
    import win32com.client
    import pythoncom
    import pywintypes
    from datetime import datetime
    
    try:
        # Step 1: Initialize
        print("\n=== Step 1: Initialize ===")
        pythoncom.CoInitialize()
        mapi.MAPIInitialize((mapi.MAPI_INIT_VERSION, mapi.MAPI_MULTITHREAD_NOTIFICATIONS))
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        session = mapi.MAPILogonEx(0, "", None, mapi.MAPI_EXTENDED | mapi.MAPI_USE_DEFAULT)
        print("✓ Initialized")
        
        # Step 2: Clean and create PST
        print("\n=== Step 2: Create PST ===")
        if pst_path.exists():
            # Remove from Outlook first
            for store in namespace.Stores:
                if store.FilePath and Path(store.FilePath).resolve() == pst_path.resolve():
                    namespace.RemoveStore(store.GetRootFolder())
                    break
            import time
            time.sleep(0.5)
            os.remove(pst_path)
            print("  Removed old PST")
        
        namespace.AddStore(str(pst_path))
        import time
        time.sleep(1)
        
        # Find our PST
        pst_store = None
        for store in namespace.Stores:
            if store.FilePath and Path(store.FilePath).resolve() == pst_path.resolve():
                pst_store = store
                break
        
        if not pst_store:
            print("❌ Could not find PST!")
            return
        print(f"✓ PST created: {pst_store.DisplayName}")
        
        # Step 3: Open via MAPI (EXACTLY like working script)
        print("\n=== Step 3: Open via MAPI ===")
        stores_table = session.GetMsgStoresTable(0)
        stores_table.SetColumns((mapitags.PR_ENTRYID, mapitags.PR_DISPLAY_NAME_A), 0)
        rows = stores_table.QueryRows(50, 0)
        
        pst_eid = None
        for row in rows:
            eid = row[0][1]
            name = row[1][1]
            if isinstance(name, bytes):
                name = name.decode('utf-8', errors='replace')
            if pst_path.stem in name or 'Minimal' in name:
                pst_eid = eid
                print(f"  Found: {name}")
                break
        
        if not pst_eid:
            # Fallback to Outlook StoreID
            pst_eid = bytes.fromhex(pst_store.StoreID)
            print("  Using Outlook StoreID")
        
        mapi_store = session.OpenMsgStore(0, pst_eid, None, mapi.MDB_WRITE | mapi.MAPI_BEST_ACCESS)
        print(f"✓ Opened store: {type(mapi_store)}")
        
        # Step 4: Get root folder (EXACTLY like working script)
        print("\n=== Step 4: Get Root Folder ===")
        PR_IPM_SUBTREE_ENTRYID = 0x35E00102
        
        props = mapi_store.GetProps([PR_IPM_SUBTREE_ENTRYID], 0)
        print(f"  GetProps returned: {props}")
        
        # Parse - working script does: props[0][1] for tuple format
        if isinstance(props[0], tuple):
            ipm_subtree_eid = props[0][1]
        elif isinstance(props, tuple) and len(props) >= 2:
            # Format: (status, ((tag, value),))
            ipm_subtree_eid = props[1][0][1]
        else:
            ipm_subtree_eid = props[0]
        
        root_folder = mapi_store.OpenEntry(ipm_subtree_eid, None, mapi.MAPI_MODIFY | mapi.MAPI_BEST_ACCESS)
        print(f"✓ Opened root: {type(root_folder)}")
        
        # Step 5: Create subfolder (EXACTLY like working script)
        print("\n=== Step 5: Create Folder ===")
        import_folder = root_folder.CreateFolder(1, "Test Folder", "Test", None, 0)
        print(f"✓ Created folder: {type(import_folder)}")
        
        # Step 6: Create message (EXACTLY like working script)
        print("\n=== Step 6: Create Message ===")
        msg = import_folder.CreateMessage(None, 0)
        print(f"✓ Created message: {type(msg)}")
        
        # Step 7: Set properties (EXACTLY like working script)
        print("\n=== Step 7: SetProps ===")
        pytime = pywintypes.Time(datetime(2020, 6, 15, 10, 30, 0))
        
        props = [
            (mapitags.PR_MESSAGE_CLASS_A, "IPM.Note"),
            (mapitags.PR_SUBJECT_A, "Test Subject"),
            (mapitags.PR_BODY_A, "Test body content"),
            (mapitags.PR_MESSAGE_FLAGS, 0x0001),  # MSGFLAG_READ
            (mapitags.PR_MESSAGE_DELIVERY_TIME, pytime),
            (mapitags.PR_CLIENT_SUBMIT_TIME, pytime),
            (mapitags.PR_SENDER_NAME_A, "Test Sender"),
            (mapitags.PR_SENDER_EMAIL_ADDRESS_A, "test@example.com"),
        ]
        
        print(f"  Setting {len(props)} properties...")
        msg.SetProps(props)
        print("✓ SetProps succeeded!")
        
        # Step 8: Save
        print("\n=== Step 8: Save ===")
        msg.SaveChanges(0)
        print("✓ Message saved!")
        
        print("\n" + "="*60)
        print("SUCCESS! Check Outlook for 'Test Folder' in the PST")
        print("="*60)
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        try:
            mapi.MAPIUninitialize()
            pythoncom.CoUninitialize()
        except:
            pass

if __name__ == '__main__':
    main()
