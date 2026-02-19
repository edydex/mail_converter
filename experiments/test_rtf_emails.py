"""Test script for analyzing RTF email conversion quality."""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logging
logging.basicConfig(level=logging.WARNING)

from core.eml_parser import EMLParser
from core.email_to_pdf import EmailToPDFConverter
from pathlib import Path

parser = EMLParser()

test_dir = '/Users/omayo/Downloads/rfp5_test_emls'
output_dir = '/Users/omayo/Downloads/rfp5_test_pdfs'
os.makedirs(output_dir, exist_ok=True)

for num in ['1', '50', '100', '135', '200', '300']:
    eml_path = f'{test_dir}/{num}.eml'
    try:
        email = parser.parse_file(eml_path)
        has_html = bool(email.body_html and email.body_html.strip())
        has_plain = bool(email.body_plain and email.body_plain.strip())
        html_len = len(email.body_html) if email.body_html else 0
        plain_len = len(email.body_plain) if email.body_plain else 0
        
        print(f'=== Email {num}: "{email.subject[:60]}" ===')
        print(f'  HTML: {has_html} ({html_len} chars), Plain: {has_plain} ({plain_len} chars)')
        print(f'  Attachments: {len(email.attachments)}, Inline images: {len(email.inline_images)}')
        
        if has_html:
            html = email.body_html
            has_div = '<div' in html.lower()
            has_br = '<br' in html.lower()
            has_b = '<b>' in html.lower() or '<b ' in html.lower()
            has_p = '<p>' in html.lower() or '<p ' in html.lower()
            has_table = '<table' in html.lower()
            has_hr = '<hr' in html.lower()
            print(f'  HTML tags: div={has_div} br={has_br} b={has_b} p={has_p} table={has_table} hr={has_hr}')
            print(f'  HTML preview: {html[:300].strip()!r}')
        
        # Save HTML for inspection
        if has_html:
            with open(f'{output_dir}/{num}_body.html', 'w') as f:
                f.write(email.body_html)
        
        # Convert to PDF
        converter = EmailToPDFConverter()
        pdf_path = Path(f'{output_dir}/{num}.pdf')
        converter.convert_email_to_pdf(email, pdf_path)
        print(f'  PDF: {pdf_path} ({pdf_path.stat().st_size} bytes)')
        
        print()
    except Exception as e:
        import traceback
        print(f'Email {num}: ERROR - {e}')
        traceback.print_exc()
        print()
