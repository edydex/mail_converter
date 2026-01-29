"""
MAPI EML Importer - Milestone 4
===============================
Goal: Import real EML files into PST with full content and preserved dates

This will:
1. Parse EML files using Python's email module
2. Create MAPI messages with all properties:
   - Subject, Body (plain + HTML)
   - From, To, CC, BCC
   - Original sent/received dates
   - Attachments
3. Save to PST file

Run with Outlook OPEN:
    python experiments/mapi_eml_import.py <path_to_eml_file_or_folder>

Example:
    python experiments/mapi_eml_import.py "C:/path/to/emails"
    python experiments/mapi_eml_import.py "C:/path/to/email.eml"
"""

import sys
import os
from pathlib import Path
from datetime import datetime
from email import message_from_bytes, policy
from email.utils import parsedate_to_datetime, getaddresses
from typing import List, Tuple, Optional
import re

def print_section(title):
    print(f"\n{'='*60}")
    print(f" {title}")
    print('='*60)


class MAPIEmlImporter:
    """
    Imports EML files into a PST using Extended MAPI.
    Preserves original sent/received dates.
    """
    
    # MAPI Property Tags
    PR_SUBJECT_W = 0x0037001F  # Unicode subject
    PR_BODY_W = 0x1000001F    # Unicode body
    PR_HTML = 0x10130102      # HTML body (binary)
    PR_MESSAGE_CLASS_W = 0x001A001F
    PR_MESSAGE_FLAGS = 0x0E070003
    PR_MESSAGE_DELIVERY_TIME = 0x0E060040
    PR_CLIENT_SUBMIT_TIME = 0x00390040
    
    # Sender properties
    PR_SENDER_NAME_W = 0x0C1A001F
    PR_SENDER_EMAIL_ADDRESS_W = 0x0C1F001F
    PR_SENT_REPRESENTING_NAME_W = 0x0042001F
    PR_SENT_REPRESENTING_EMAIL_ADDRESS_W = 0x0065001F
    PR_SENDER_ADDRTYPE_W = 0x0C1E001F
    PR_SENT_REPRESENTING_ADDRTYPE_W = 0x0064001F
    
    # Display To/CC
    PR_DISPLAY_TO_W = 0x0E04001F
    PR_DISPLAY_CC_W = 0x0E03001F
    PR_DISPLAY_BCC_W = 0x0E02001F
    
    # Message flags
    MSGFLAG_READ = 0x0001
    MSGFLAG_UNSENT = 0x0008
    
    # Recipient types
    MAPI_TO = 1
    MAPI_CC = 2
    MAPI_BCC = 3
    
    def __init__(self):
        self.mapi = None
        self.mapitags = None
        self.session = None
        self.outlook = None
        self.namespace = None
        self.pythoncom = None
        self.pywintypes = None
        
    def initialize(self) -> bool:
        """Initialize MAPI and connect to Outlook."""
        try:
            from win32com.mapi import mapi, mapitags
            import win32com.client
            import pythoncom
            import pywintypes
            
            self.mapi = mapi
            self.mapitags = mapitags
            self.pythoncom = pythoncom
            self.pywintypes = pywintypes
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Initialize MAPI
            mapi.MAPIInitialize((mapi.MAPI_INIT_VERSION, mapi.MAPI_MULTITHREAD_NOTIFICATIONS))
            
            # Connect to Outlook
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Get MAPI session
            self.session = mapi.MAPILogonEx(0, "", None, mapi.MAPI_EXTENDED | mapi.MAPI_USE_DEFAULT)
            
            return True
            
        except Exception as e:
            print(f"❌ Initialization failed: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def cleanup(self):
        """Clean up MAPI resources."""
        try:
            if self.mapi:
                self.mapi.MAPIUninitialize()
            if self.pythoncom:
                self.pythoncom.CoUninitialize()
        except:
            pass
    
    def create_or_open_pst(self, pst_path: str) -> Tuple[object, object]:
        """
        Create a new PST or open existing one.
        Returns (mapi_store, outlook_store) tuple.
        """
        pst_path = Path(pst_path).resolve()
        
        # Check if PST already mounted
        outlook_store = None
        for store in self.namespace.Stores:
            if store.FilePath:
                if Path(store.FilePath).resolve() == pst_path:
                    outlook_store = store
                    print(f"✓ PST already mounted: {store.DisplayName}")
                    break
        
        # Add to profile if not mounted
        if not outlook_store:
            print(f"Adding PST to profile: {pst_path}")
            self.namespace.AddStore(str(pst_path))
            
            # Find it
            import time
            time.sleep(0.5)
            for store in self.namespace.Stores:
                if store.FilePath:
                    if Path(store.FilePath).resolve() == pst_path:
                        outlook_store = store
                        print(f"✓ PST added: {store.DisplayName}")
                        break
        
        if not outlook_store:
            raise RuntimeError(f"Could not find PST store: {pst_path}")
        
        # Open via MAPI
        store_id = bytes.fromhex(outlook_store.StoreID)
        mapi_store = self.session.OpenMsgStore(
            0, store_id, None, 
            self.mapi.MDB_WRITE | self.mapi.MAPI_BEST_ACCESS
        )
        
        return mapi_store, outlook_store
    
    def get_or_create_folder(self, mapi_store, outlook_store, folder_name: str):
        """Get or create a folder in the PST using pure MAPI (same as working script)."""
        import time
        
        # Get IPM Subtree - this is where mail folders live (same as working mapi_pst_create.py)
        PR_IPM_SUBTREE_ENTRYID = 0x35E00102
        result = mapi_store.GetProps([PR_IPM_SUBTREE_ENTRYID], 0)
        
        # Parse result - handle different formats
        if isinstance(result, list) and len(result) > 0:
            if isinstance(result[0], tuple):
                root_eid = result[0][1]
            else:
                root_eid = result[0]
        else:
            raise RuntimeError("Could not get IPM Subtree entry ID")
        
        # Open root folder with full access
        root_folder = mapi_store.OpenEntry(
            root_eid, None, 
            self.mapi.MAPI_MODIFY | self.mapi.MAPI_BEST_ACCESS
        )
        
        # Try to create the subfolder
        PR_ENTRYID = 0x0FFF0102
        try:
            new_folder = root_folder.CreateFolder(1, folder_name, "", None, 0)
            print(f"✓ Created folder: {folder_name}")
            
            # Get its entry ID
            folder_props = new_folder.GetProps([PR_ENTRYID], 0)
            if isinstance(folder_props, list) and len(folder_props) > 0:
                if isinstance(folder_props[0], tuple):
                    folder_eid = folder_props[0][1]
                else:
                    folder_eid = folder_props[0]
            else:
                raise RuntimeError("Could not get new folder entry ID")
                
        except Exception as e:
            if "MAPI_E_COLLISION" in str(e) or "already exists" in str(e).lower():
                print(f"Folder exists, searching...")
                # Folder exists - find it via MAPI table
                table = root_folder.GetHierarchyTable(0)
                table.SetColumns([PR_ENTRYID, 0x3001001F], 0)  # PR_ENTRYID, PR_DISPLAY_NAME_W
                
                folder_eid = None
                while True:
                    rows = table.QueryRows(10, 0)
                    if not rows:
                        break
                    for row in rows:
                        eid, name = row[0][1], row[1][1]
                        if name == folder_name:
                            folder_eid = eid
                            break
                    if folder_eid:
                        break
                
                if not folder_eid:
                    raise RuntimeError(f"Could not find existing folder: {folder_name}")
            else:
                raise
        
        # Re-open folder with full write access
        folder = mapi_store.OpenEntry(
            folder_eid, None,
            self.mapi.MAPI_MODIFY | self.mapi.MAPI_BEST_ACCESS
        )
        print(f"✓ Opened folder with write access")
        return folder
    
    def parse_eml(self, eml_path: str) -> dict:
        """Parse an EML file and extract all components."""
        with open(eml_path, 'rb') as f:
            msg = message_from_bytes(f.read(), policy=policy.default)
        
        result = {
            'subject': msg.get('Subject', '(No Subject)') or '(No Subject)',
            'from_name': '',
            'from_email': '',
            'to': [],      # List of (name, email)
            'cc': [],
            'bcc': [],
            'date': None,
            'body_plain': '',
            'body_html': '',
            'attachments': [],  # List of (filename, content_type, data)
        }
        
        # Parse From
        from_header = msg.get('From', '')
        if from_header:
            addrs = getaddresses([from_header])
            if addrs:
                result['from_name'] = addrs[0][0] or addrs[0][1]
                result['from_email'] = addrs[0][1]
        
        # Parse To, CC, BCC
        for field, key in [('To', 'to'), ('Cc', 'cc'), ('Bcc', 'bcc')]:
            header = msg.get(field, '')
            if header:
                result[key] = getaddresses([header])
        
        # Parse Date
        date_str = msg.get('Date', '')
        if date_str:
            try:
                result['date'] = parsedate_to_datetime(date_str)
            except:
                result['date'] = datetime.now()
        else:
            result['date'] = datetime.now()
        
        # Parse Body
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                disposition = part.get('Content-Disposition', '')
                
                # Skip attachments for body extraction
                if 'attachment' in disposition:
                    continue
                
                if content_type == 'text/plain' and not result['body_plain']:
                    try:
                        result['body_plain'] = part.get_content()
                    except:
                        pass
                        
                elif content_type == 'text/html' and not result['body_html']:
                    try:
                        result['body_html'] = part.get_content()
                    except:
                        pass
            
            # Extract attachments
            for part in msg.iter_attachments():
                filename = part.get_filename()
                if filename:
                    try:
                        data = part.get_payload(decode=True)
                        if data:
                            result['attachments'].append({
                                'filename': filename,
                                'content_type': part.get_content_type(),
                                'data': data,
                            })
                    except:
                        pass
        else:
            # Single part message
            content_type = msg.get_content_type()
            try:
                content = msg.get_content()
                if content_type == 'text/html':
                    result['body_html'] = content
                else:
                    result['body_plain'] = content
            except:
                result['body_plain'] = msg.get_payload(decode=True).decode('utf-8', errors='replace')
        
        return result
    
    def import_eml(self, folder, eml_data: dict) -> bool:
        """Import parsed EML data into a MAPI folder."""
        try:
            # Create message
            msg = folder.CreateMessage(None, 0)
            
            # Convert date to PyTime
            pytime = self.pywintypes.Time(eml_data['date'])
            
            # Build properties list
            props = [
                (self.PR_MESSAGE_CLASS_W, "IPM.Note"),
                (self.PR_SUBJECT_W, eml_data['subject']),
                (self.PR_MESSAGE_FLAGS, self.MSGFLAG_READ),  # Mark as read, not unsent
                (self.PR_MESSAGE_DELIVERY_TIME, pytime),
                (self.PR_CLIENT_SUBMIT_TIME, pytime),
            ]
            
            # Sender
            if eml_data['from_email']:
                props.extend([
                    (self.PR_SENDER_NAME_W, eml_data['from_name'] or eml_data['from_email']),
                    (self.PR_SENDER_EMAIL_ADDRESS_W, eml_data['from_email']),
                    (self.PR_SENDER_ADDRTYPE_W, "SMTP"),
                    (self.PR_SENT_REPRESENTING_NAME_W, eml_data['from_name'] or eml_data['from_email']),
                    (self.PR_SENT_REPRESENTING_EMAIL_ADDRESS_W, eml_data['from_email']),
                    (self.PR_SENT_REPRESENTING_ADDRTYPE_W, "SMTP"),
                ])
            
            # Body - prefer HTML if available
            if eml_data['body_html']:
                # HTML body needs to be bytes
                html_bytes = eml_data['body_html'].encode('utf-8')
                props.append((self.PR_HTML, html_bytes))
                # Also set plain text version
                if eml_data['body_plain']:
                    props.append((self.PR_BODY_W, eml_data['body_plain']))
            elif eml_data['body_plain']:
                props.append((self.PR_BODY_W, eml_data['body_plain']))
            
            # Display To/CC
            if eml_data['to']:
                display_to = '; '.join([f"{n} <{e}>" if n else e for n, e in eml_data['to']])
                props.append((self.PR_DISPLAY_TO_W, display_to))
            
            if eml_data['cc']:
                display_cc = '; '.join([f"{n} <{e}>" if n else e for n, e in eml_data['cc']])
                props.append((self.PR_DISPLAY_CC_W, display_cc))
            
            # Set properties
            msg.SetProps(props)
            
            # Add recipients
            self._add_recipients(msg, eml_data)
            
            # Add attachments
            for att in eml_data['attachments']:
                self._add_attachment(msg, att)
            
            # Save
            msg.SaveChanges(0)
            
            return True
            
        except Exception as e:
            print(f"❌ Error importing email: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _add_recipients(self, msg, eml_data: dict):
        """Add recipients to the message."""
        # Recipient property tags
        PR_RECIPIENT_TYPE = 0x0C150003
        PR_DISPLAY_NAME_W = 0x3001001F
        PR_EMAIL_ADDRESS_W = 0x3003001F
        PR_ADDRTYPE_W = 0x3002001F
        PR_ENTRYID = 0x0FFF0102
        
        recipients = []
        
        for recip_type, recip_list in [
            (self.MAPI_TO, eml_data['to']),
            (self.MAPI_CC, eml_data['cc']),
            (self.MAPI_BCC, eml_data['bcc']),
        ]:
            for name, email in recip_list:
                recipients.append({
                    'type': recip_type,
                    'name': name or email,
                    'email': email,
                })
        
        if not recipients:
            return
        
        try:
            # Build ADRLIST structure
            # Each recipient is a row with properties
            adrlist = []
            for recip in recipients:
                row = [
                    (PR_RECIPIENT_TYPE, recip['type']),
                    (PR_DISPLAY_NAME_W, recip['name']),
                    (PR_EMAIL_ADDRESS_W, recip['email']),
                    (PR_ADDRTYPE_W, "SMTP"),
                ]
                adrlist.append(row)
            
            # ModifyRecipients(flags, adrlist)
            # flags: MODRECIP_ADD = 2
            msg.ModifyRecipients(2, adrlist)
            
        except Exception as e:
            # Recipients are optional - continue without them
            print(f"  ⚠ Could not add recipients: {e}")
    
    def _add_attachment(self, msg, attachment: dict):
        """Add an attachment to the message."""
        PR_ATTACH_METHOD = 0x37050003
        PR_ATTACH_FILENAME_W = 0x3704001F
        PR_ATTACH_LONG_FILENAME_W = 0x3707001F
        PR_ATTACH_DATA_BIN = 0x37010102
        PR_ATTACH_MIME_TAG_W = 0x370E001F
        PR_DISPLAY_NAME_W = 0x3001001F
        
        ATTACH_BY_VALUE = 1
        
        try:
            # Create attachment
            attach_num, attach = msg.CreateAttach(None, 0)
            
            props = [
                (PR_ATTACH_METHOD, ATTACH_BY_VALUE),
                (PR_ATTACH_FILENAME_W, attachment['filename'][:255]),
                (PR_ATTACH_LONG_FILENAME_W, attachment['filename']),
                (PR_ATTACH_DATA_BIN, attachment['data']),
                (PR_DISPLAY_NAME_W, attachment['filename']),
            ]
            
            if attachment.get('content_type'):
                props.append((PR_ATTACH_MIME_TAG_W, attachment['content_type']))
            
            attach.SetProps(props)
            attach.SaveChanges(0)
            
        except Exception as e:
            print(f"  ⚠ Could not add attachment '{attachment['filename']}': {e}")


def main():
    print("MAPI EML Importer - Milestone 4")
    
    if sys.platform != 'win32':
        print("❌ This script must be run on Windows!")
        return
    
    # Get input path from command line or use test
    if len(sys.argv) > 1:
        input_path = Path(sys.argv[1])
    else:
        # Create a test EML for demonstration
        input_path = None
    
    # Output PST
    documents = Path(os.environ.get('USERPROFILE', '')) / 'Documents'
    pst_path = documents / 'MAPI_Import_Test.pst'
    
    print(f"Output PST: {pst_path}")
    
    # Initialize importer
    importer = MAPIEmlImporter()
    
    try:
        print_section("Step 1: Initialize MAPI")
        if not importer.initialize():
            return
        print("✓ MAPI initialized")
        
        print_section("Step 2: Create/Open PST")
        mapi_store, outlook_store = importer.create_or_open_pst(str(pst_path))
        print("✓ PST ready")
        
        print_section("Step 3: Create Import Folder")
        folder = importer.get_or_create_folder(mapi_store, outlook_store, "EML Imports")
        print("✓ Folder ready")
        
        print_section("Step 4: Import EMLs")
        
        if input_path and input_path.exists():
            # Import from specified path
            eml_files = []
            if input_path.is_file() and input_path.suffix.lower() == '.eml':
                eml_files = [input_path]
            elif input_path.is_dir():
                eml_files = list(input_path.glob('*.eml')) + list(input_path.glob('*.EML'))
                # Also check for files without extension (readpst output)
                for f in input_path.iterdir():
                    if f.is_file() and not f.suffix:
                        # Check if it looks like an email
                        try:
                            with open(f, 'rb') as test_f:
                                start = test_f.read(100)
                                if b'From:' in start or b'Subject:' in start or b'Date:' in start:
                                    eml_files.append(f)
                        except:
                            pass
            
            print(f"Found {len(eml_files)} EML files to import")
            
            success_count = 0
            for i, eml_file in enumerate(eml_files):
                print(f"\n[{i+1}/{len(eml_files)}] {eml_file.name}")
                
                try:
                    eml_data = importer.parse_eml(str(eml_file))
                    print(f"  Subject: {eml_data['subject'][:50]}...")
                    print(f"  Date: {eml_data['date']}")
                    print(f"  From: {eml_data['from_email']}")
                    print(f"  Attachments: {len(eml_data['attachments'])}")
                    
                    if importer.import_eml(folder, eml_data):
                        print(f"  ✓ Imported successfully")
                        success_count += 1
                    else:
                        print(f"  ❌ Import failed")
                        
                except Exception as e:
                    print(f"  ❌ Error: {e}")
            
            print(f"\n✓ Imported {success_count}/{len(eml_files)} emails")
            
        else:
            # Create test emails for demonstration
            print("No input path specified - creating test emails")
            
            test_emails = [
                {
                    'subject': 'Test Email with Attachment',
                    'from_name': 'Test Sender',
                    'from_email': 'sender@example.com',
                    'to': [('Recipient One', 'recipient@example.com')],
                    'cc': [('CC Person', 'cc@example.com')],
                    'bcc': [],
                    'date': datetime(2019, 12, 25, 10, 0, 0),
                    'body_plain': 'This is a test email with plain text body.\n\nIt has multiple lines.',
                    'body_html': '<html><body><h1>Test Email</h1><p>This is <b>HTML</b> content.</p></body></html>',
                    'attachments': [
                        {
                            'filename': 'test.txt',
                            'content_type': 'text/plain',
                            'data': b'This is a test attachment file content.',
                        }
                    ],
                },
                {
                    'subject': 'Another Historical Email',
                    'from_name': 'John Doe',
                    'from_email': 'john.doe@company.com',
                    'to': [('Jane Smith', 'jane@company.com'), ('Bob Wilson', 'bob@company.com')],
                    'cc': [],
                    'bcc': [],
                    'date': datetime(2017, 3, 14, 15, 30, 0),
                    'body_plain': 'Meeting reminder for tomorrow at 3 PM.',
                    'body_html': '',
                    'attachments': [],
                },
            ]
            
            for i, eml_data in enumerate(test_emails):
                print(f"\n[{i+1}/{len(test_emails)}] {eml_data['subject']}")
                print(f"  Date: {eml_data['date']}")
                
                if importer.import_eml(folder, eml_data):
                    print(f"  ✓ Created successfully")
                else:
                    print(f"  ❌ Creation failed")
        
        print_section("SUCCESS!")
        print(f"""
Emails imported to: {pst_path}
Folder: "EML Imports"

CHECK IN OUTLOOK:
1. Find the PST in your folder list
2. Open "EML Imports" folder
3. Verify:
   - Email dates are ORIGINAL dates, not today
   - Sender/recipient info is correct
   - HTML body renders properly
   - Attachments are present
""")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        print_section("Cleanup")
        importer.cleanup()
        print("✓ Done")


if __name__ == '__main__':
    main()
