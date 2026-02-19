"""Precisely trace what happens to the first font-family inline style."""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logging
logging.basicConfig(level=logging.ERROR)

from core.eml_parser import EMLParser

parser = EMLParser()
email = parser.parse_file('/Users/omayo/Downloads/rfp5_test_emls/1.eml')

raw = email.body_html

# Find the first inline style with font-family
idx = raw.find('style="font-family:"Aptos"')
if idx < 0:
    idx = raw.find('style="font-family:"')
print(f"First inline font-family at position: {idx}")
print(f"Context (raw): {raw[idx:idx+100]!r}")

# Apply the font-family fix  
html = raw
_prev = None
iterations = 0
while _prev != html:
    _prev = html
    html = re.sub(
        r'(font-family\s*:\s*[^";>]*?)"([^"]*?)"',
        r"\1'\2'",
        html,
    )
    iterations += 1
print(f"\nFont-family fix iterations: {iterations}")

# Check the same position after fix
print(f"\nContext (after fix): {html[idx:idx+100]!r}")

# Does the fix result in the expected pattern?
expected = '''style="font-family:'Aptos',sans-serif"'''
found = expected in html
print(f"\nExpected pattern '{expected}' found: {found}")

# What IS actually at that position?
# Find what style="font-family: looks like after fix
for m in re.finditer(r"style=\"font-family:", html):
    pos = m.start()
    chunk = html[pos:pos+60]
    print(f"  At pos {pos}: {chunk!r}")
    if pos > idx + 200:
        break

# Check: does 'sans-serif">' exist (properly closed style)?
count_proper = html.count("sans-serif\">")
count_broken = html.count("sans-serif'>")
print(f"\nProperly closed 'sans-serif\">': {count_proper}")
print(f"Broken 'sans-serif\\'>': {count_broken}")

# Also check for sans-serif;mso pattern
count_mso_proper = len(re.findall(r'sans-serif;[^"]*">', html))
print(f"Pattern 'sans-serif;...\">' (proper with mso): {count_mso_proper}")

# Show what comes right after font-family:'Aptos',sans-serif
for m in re.finditer(r"font-family:'Aptos',sans-serif", html):
    pos = m.end()
    after = html[pos:pos+5]
    print(f"  After 'font-family:'Aptos',sans-serif': {after!r}")
