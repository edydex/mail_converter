"""Analyze what kind of RTF each email has and what the de-encapsulation produces."""
import sys, os, base64, email
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

test_dir = '/Users/omayo/Downloads/rfp5_test_emls'

for num in ['1', '135', '300']:
    eml_path = f'{test_dir}/{num}.eml'
    with open(eml_path, 'rb') as f:
        msg = email.message_from_bytes(f.read())
    
    print(f"=== Email {num}: {msg['Subject'][:60]} ===")
    
    for part in msg.walk():
        ct = part.get_content_type()
        fn = part.get_filename()
        payload = part.get_payload(decode=True)
        if ct == 'application/rtf' or (fn and 'rtf' in fn.lower()):
            print(f"  RTF part: content_type={ct}, filename={fn}, size={len(payload)} bytes")
            
            # Check for \fromhtml marker
            rtf_text = payload[:2000].decode('ascii', errors='replace')
            has_fromhtml = '\\fromhtml' in rtf_text
            has_fromtext = '\\fromtext' in rtf_text
            print(f"  \\fromhtml: {has_fromhtml}")
            print(f"  \\fromtext: {has_fromtext}")
            
            # Show first 500 bytes
            snippet = payload[:500].decode('ascii', errors='replace')
            # Truncate long lines
            lines = snippet.split('\n')
            for line in lines[:10]:
                print(f"  | {line[:120]}")
            
            # Check for interesting RTF control words
            from core.rtf_converter import convert_rtf_body
            plain, html = convert_rtf_body(payload)
            print(f"\n  Converted: plain={len(plain)} chars, html={len(html)} chars")
            if html:
                # Check for separator patterns
                has_hr = '<hr' in html.lower()
                has_separator = '-----' in html or '___' in html
                has_bold = '<b>' in html.lower() or 'font-weight' in html.lower()
                has_highlight = 'background' in html.lower() or 'highlight' in html.lower()
                print(f"  HTML has <hr>: {has_hr}")
                print(f"  HTML has text separator: {has_separator}")
                print(f"  HTML has bold: {has_bold}")
                print(f"  HTML has highlight/background: {has_highlight}")
        
        elif ct == 'text/plain':
            if payload:
                text = payload.decode('utf-8', errors='replace')
                print(f"  text/plain part: {len(text)} chars")
        elif ct == 'text/html':
            if payload:
                text = payload.decode('utf-8', errors='replace')
                print(f"  text/html part: {len(text)} chars")
    
    print()
