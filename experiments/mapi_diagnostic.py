"""
DIAGNOSTIC TEST: Compare working pattern with EML import
========================================================
This test does two things:
1. Creates a test message the EXACT same way as mapi_pst_create.py (should work)
2. Creates a message from an EML file (might fail)

This helps isolate where the problem is.
"""

import sys
import os
from pathlib import Path

def main():
    print("MAPI Diagnostic Test")
    print("=" * 60)
    
    if sys.platform != 'win32':
        print("❌ Windows only!")
        return
    
    # Get EML path from command line
    eml_path = sys.argv[1] if len(sys.argv) > 1 else None
    
    # Output PST
    documents = Path(os.environ.get('USERPROFILE', '')) / 'Documents'
    from datetime import datetime as dt
    pst_path = documents / f'MAPI_Diagnostic_{dt.now().strftime("%Y%m%d_%H%M%S")}.pst'
    
    print(f"Target PST: {pst_path}")
    if eml_path:
        print(f"Test EML: {eml_path}")
    
    # Import modules - EXACTLY like working script
    from win32com.mapi import mapi, mapitags
    import win32com.client
    import pythoncom
    import pywintypes
    from datetime import datetime
    
    try:
        # Step 1: Initialize - EXACTLY like working script
        print("\n=== Step 1: Initialize ===")
        pythoncom.CoInitialize()
        mapi.MAPIInitialize((mapi.MAPI_INIT_VERSION, mapi.MAPI_MULTITHREAD_NOTIFICATIONS))
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        session = mapi.MAPILogonEx(0, "", None, mapi.MAPI_EXTENDED | mapi.MAPI_USE_DEFAULT)
        print("✓ Initialized")
        
        # Step 2: Create PST - EXACTLY like working script
        print("\n=== Step 2: Create PST ===")
        namespace.AddStore(str(pst_path))
        import time
        time.sleep(1)
        
        pst_store = None
        for store in namespace.Stores:
            if store.FilePath and Path(store.FilePath).resolve() == pst_path.resolve():
                pst_store = store
                break
        
        if not pst_store:
            print("❌ Could not find PST!")
            return
        print(f"✓ PST created: {pst_store.DisplayName}")
        
        # Step 3: Open via MAPI - EXACTLY like working script
        print("\n=== Step 3: Open via MAPI ===")
        pst_eid = bytes.fromhex(pst_store.StoreID)
        mapi_store = session.OpenMsgStore(0, pst_eid, None, mapi.MDB_WRITE | mapi.MAPI_BEST_ACCESS)
        print(f"✓ Opened store")
        
        # Step 4: Get root folder - EXACTLY like working script
        print("\n=== Step 4: Get Root Folder ===")
        PR_IPM_SUBTREE_ENTRYID = 0x35E00102
        props = mapi_store.GetProps([PR_IPM_SUBTREE_ENTRYID], 0)
        
        if isinstance(props[0], tuple):
            ipm_subtree_eid = props[0][1]
        elif isinstance(props, tuple) and len(props) >= 2:
            ipm_subtree_eid = props[1][0][1]
        else:
            ipm_subtree_eid = props[0]
        
        root_folder = mapi_store.OpenEntry(ipm_subtree_eid, None, mapi.MAPI_MODIFY | mapi.MAPI_BEST_ACCESS)
        print(f"✓ Opened root folder")
        
        # Step 5: Create folder - EXACTLY like working script
        print("\n=== Step 5: Create Folder ===")
        import_folder = root_folder.CreateFolder(1, "Diagnostic Test", "Test", None, 0)
        print(f"✓ Created folder")
        
        # =====================================================================
        # TEST 1: Create message EXACTLY like working script
        # =====================================================================
        print("\n" + "=" * 60)
        print("TEST 1: Create test message (exact working pattern)")
        print("=" * 60)
        
        PR_MESSAGE_CLASS = mapitags.PR_MESSAGE_CLASS_A
        PR_SUBJECT = mapitags.PR_SUBJECT_A
        PR_BODY = mapitags.PR_BODY_A
        PR_MESSAGE_FLAGS = mapitags.PR_MESSAGE_FLAGS
        PR_MESSAGE_DELIVERY_TIME = mapitags.PR_MESSAGE_DELIVERY_TIME
        PR_CLIENT_SUBMIT_TIME = mapitags.PR_CLIENT_SUBMIT_TIME
        PR_SENDER_NAME = mapitags.PR_SENDER_NAME_A
        PR_SENDER_EMAIL = mapitags.PR_SENDER_EMAIL_ADDRESS_A
        PR_SENT_REP_NAME = 0x0042001E
        PR_SENT_REP_EMAIL = 0x0065001E
        MSGFLAG_READ = 0x0001
        
        try:
            msg = import_folder.CreateMessage(None, 0)
            print(f"  ✓ CreateMessage succeeded: {type(msg)}")
            
            pytime = pywintypes.Time(datetime(2020, 6, 15, 10, 30, 0))
            
            props = [
                (PR_MESSAGE_CLASS, "IPM.Note"),
                (PR_SUBJECT, "Test Message - Hardcoded"),
                (PR_BODY, "This is a test body"),
                (PR_MESSAGE_FLAGS, MSGFLAG_READ),
                (PR_MESSAGE_DELIVERY_TIME, pytime),
                (PR_CLIENT_SUBMIT_TIME, pytime),
                (PR_SENDER_NAME, "Test Sender"),
                (PR_SENDER_EMAIL, "test@example.com"),
                (PR_SENT_REP_NAME, "Test Sender"),
                (PR_SENT_REP_EMAIL, "test@example.com"),
            ]
            
            print(f"  Setting {len(props)} properties...")
            for i, (tag, val) in enumerate(props):
                print(f"    [{i}] Tag: 0x{tag:08X}, Value type: {type(val).__name__}, Value: {repr(val)[:50]}")
            
            msg.SetProps(props)
            print("  ✓ SetProps succeeded!")
            
            msg.SaveChanges(0)
            print("✓ TEST 1 PASSED: Hardcoded message created!")
            
        except Exception as e:
            print(f"❌ TEST 1 FAILED: {e}")
            import traceback
            traceback.print_exc()
        
        # =====================================================================
        # TEST 2: Create message from EML data
        # =====================================================================
        if eml_path and Path(eml_path).exists():
            print("\n" + "=" * 60)
            print("TEST 2: Create message from EML")
            print("=" * 60)
            
            # Parse EML
            from email import message_from_bytes, policy
            from email.utils import parsedate_to_datetime
            
            with open(eml_path, 'rb') as f:
                eml = message_from_bytes(f.read(), policy=policy.default)
            
            subject = eml.get('Subject', 'No Subject') or 'No Subject'
            from_header = eml.get('From', '')
            date_str = eml.get('Date', '')
            
            print(f"  EML Subject: {subject[:50]}")
            print(f"  EML From: {from_header}")
            print(f"  EML Date: {date_str}")
            
            # Parse date
            try:
                email_date = parsedate_to_datetime(date_str)
            except:
                email_date = datetime.now()
            
            # Get body
            body = ""
            if eml.is_multipart():
                for part in eml.walk():
                    if part.get_content_type() == 'text/plain':
                        try:
                            body = part.get_content()
                            break
                        except:
                            pass
            else:
                try:
                    body = eml.get_content()
                except:
                    body = str(eml.get_payload(decode=True))
            
            # Limit body length
            body = body[:1000] if body else "No body"
            
            # Encode for ANSI
            def safe_ansi(s):
                if not s:
                    return ""
                return s.encode('latin-1', errors='replace').decode('latin-1')
            
            try:
                msg2 = import_folder.CreateMessage(None, 0)
                print(f"  ✓ CreateMessage succeeded: {type(msg2)}")
                
                pytime2 = pywintypes.Time(email_date.replace(tzinfo=None))
                
                props2 = [
                    (PR_MESSAGE_CLASS, "IPM.Note"),
                    (PR_SUBJECT, safe_ansi(subject[:255])),
                    (PR_BODY, safe_ansi(body)),
                    (PR_MESSAGE_FLAGS, MSGFLAG_READ),
                    (PR_MESSAGE_DELIVERY_TIME, pytime2),
                    (PR_CLIENT_SUBMIT_TIME, pytime2),
                    (PR_SENDER_NAME, safe_ansi(from_header[:100])),
                    (PR_SENDER_EMAIL, safe_ansi(from_header[:100])),
                ]
                
                print(f"  Setting {len(props2)} properties...")
                for i, (tag, val) in enumerate(props2):
                    val_repr = repr(val)[:50] if not isinstance(val, pywintypes.TimeType) else str(val)
                    print(f"    [{i}] Tag: 0x{tag:08X}, Type: {type(val).__name__}, Val: {val_repr}")
                
                msg2.SetProps(props2)
                print("  ✓ SetProps succeeded!")
                
                msg2.SaveChanges(0)
                print("✓ TEST 2 PASSED: EML message created!")
                
            except Exception as e:
                print(f"❌ TEST 2 FAILED: {e}")
                import traceback
                traceback.print_exc()
        
        # =====================================================================
        # TEST 3: Minimal props - just subject and class
        # =====================================================================
        print("\n" + "=" * 60)
        print("TEST 3: Minimal properties (just message class and subject)")
        print("=" * 60)
        
        try:
            msg3 = import_folder.CreateMessage(None, 0)
            print(f"  ✓ CreateMessage succeeded")
            
            props3 = [
                (PR_MESSAGE_CLASS, "IPM.Note"),
                (PR_SUBJECT, "Minimal Test"),
            ]
            
            print(f"  Setting {len(props3)} properties...")
            msg3.SetProps(props3)
            print("  ✓ SetProps succeeded!")
            
            msg3.SaveChanges(0)
            print("✓ TEST 3 PASSED!")
            
        except Exception as e:
            print(f"❌ TEST 3 FAILED: {e}")
            import traceback
            traceback.print_exc()
        
        print("\n" + "=" * 60)
        print("DIAGNOSTIC COMPLETE")
        print("=" * 60)
        print(f"\nCheck Outlook for PST: {pst_path}")
        
    except Exception as e:
        print(f"\n❌ FATAL ERROR: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        try:
            mapi.MAPIUninitialize()
            pythoncom.CoUninitialize()
        except:
            pass

if __name__ == "__main__":
    main()
