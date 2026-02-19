"""Prove that BeautifulSoup in _strip_fixed_table_widths corrupts HTML."""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logging
logging.basicConfig(level=logging.ERROR)

from core.eml_parser import EMLParser
from core.email_to_pdf import EmailToPDFConverter

parser = EMLParser()
email = parser.parse_file('/Users/omayo/Downloads/rfp5_test_emls/1.eml')

converter = EmailToPDFConverter()
sanitized = converter._sanitize_email_html(email.body_html)

# Count corruption BEFORE BeautifulSoup
before = len(re.findall(r'style="[^"]*&gt;', sanitized))
print(f'Corrupted style attrs BEFORE _strip_fixed_table_widths: {before}')

# Run through _strip_fixed_table_widths (which uses BeautifulSoup)
after_bs = converter._strip_fixed_table_widths(sanitized)

# Count corruption AFTER BeautifulSoup
after = len(re.findall(r'style="[^"]*&gt;', after_bs))
print(f'Corrupted style attrs AFTER _strip_fixed_table_widths: {after}')

if after > 0:
    print(f'\n  ==> CONFIRMED: BeautifulSoup is corrupting the HTML!\n')
    # Show examples
    for m in re.finditer(r'style="[^"]*&gt;[^"]*"', after_bs):
        print(f'  EXAMPLE: {m.group()[:200]}')
        print()
        break
else:
    print('\n  BeautifulSoup is NOT the problem. Need to investigate further.')

# Also check: what does the raw HTML look like before sanitization?
raw = email.body_html
raw_broken = re.findall(r'style="[^"]*font-family\s*:\s*"', raw)
print(f'\nBroken style attrs in RAW HTML (before fix): {len(raw_broken)}')

# Check if sanitization really fixed them
san_broken = re.findall(r'style="[^"]*font-family\s*:\s*"', sanitized)
print(f'Broken style attrs in SANITIZED HTML (after fix): {len(san_broken)}')

# Check the _clean_inline_mso effect
# Count style attrs with font-family in raw vs sanitized
raw_ff_styles = re.findall(r'style="[^"]*font-family[^"]*"', raw)
san_ff_styles = re.findall(r'style="[^"]*font-family[^"]*"', sanitized)
print(f'\nStyle attrs with font-family:')
print(f'  In raw HTML: {len(raw_ff_styles)}')
print(f'  In sanitized HTML: {len(san_ff_styles)}')

# Show raw examples
print(f'\nRaw style attr examples with font-family:')
for s in raw_ff_styles[:3]:
    print(f'  {s[:150]}')
print(f'\nSanitized style attr examples with font-family:')
for s in san_ff_styles[:3]:
    print(f'  {s[:150]}')
