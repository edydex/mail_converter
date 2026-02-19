"""Examine the exact raw HTML structure around font-family patterns."""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logging
logging.basicConfig(level=logging.ERROR)

from core.eml_parser import EMLParser

parser = EMLParser()
email = parser.parse_file('/Users/omayo/Downloads/rfp5_test_emls/1.eml')

raw = email.body_html

# Find all occurrences of font-family: in inline styles (not CSS blocks)
# Look for style=" ... font-family: ... " pattern
for m in re.finditer(r'style="', raw):
    start = m.start()
    # Find the raw HTML from style=" to the next few "
    after = raw[start:start+300]
    
    # Only show if it contains font-family
    if 'font-family' in after and 'font-family:' in after:
        # Show the style attribute and what follows
        print(f"Position {start}:")
        print(f"  Raw: {after[:200]!r}")
        
        # Count " chars in the first 100 characters
        first100 = raw[start+7:start+200]  # skip 'style="'
        quotes = [i for i, c in enumerate(first100) if c == '"']
        print(f'  " positions (relative to style="): {quotes[:10]}')
        print()
