"""Trace the exact point of HTML corruption in _sanitize_email_html."""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logging
logging.basicConfig(level=logging.ERROR)

from core.eml_parser import EMLParser
from core.email_to_pdf import EmailToPDFConverter

parser = EMLParser()
email = parser.parse_file('/Users/omayo/Downloads/rfp5_test_emls/1.eml')

raw = email.body_html

# Step 1: Apply just the font-family fix (same as _sanitize_email_html step 3)
html = raw
_prev = None
while _prev != html:
    _prev = html
    html = re.sub(
        r'(font-family\s*:\s*[^";>]*?)"([^"]*?)"',
        r"\1'\2'",
        html,
    )
html_after_ff_fix = html

# Check: is the HTML correct at this point?
# Look for style attrs that DON'T properly close
# A proper style attr is style="...content-without-double-quotes..."
# A broken one would have style="...short..." followed by more stuff before next >
broken_check = re.findall(r'style="[^"]*font-family[^"]*"', html_after_ff_fix)
print(f"After font-family fix - style attrs with font-family: {len(broken_check)}")
for bc in broken_check[:3]:
    if len(bc) > 200:
        print(f"  LONG MATCH ({len(bc)} chars): {bc[:100]}...{bc[-50:]}")
    else:
        print(f"  {bc}")

# Step 2: Apply _clean_inline_mso (same as _sanitize_email_html step 9)
def _clean_inline_mso(m):
    style = m.group(1)
    style = re.sub(r'\bmso-[a-z\-]+\s*:[^;"]+;?\s*', '', style, flags=re.IGNORECASE)
    style = re.sub(r'\btab-stops\s*:[^;"]+;?\s*', '', style, flags=re.IGNORECASE)
    style = style.strip()
    if not style:
        return ''
    return f'style="{style}"'

html_after_mso_clean = re.sub(r'style="([^"]*)"', _clean_inline_mso, html_after_ff_fix, flags=re.IGNORECASE)

# Check again
broken_check2 = re.findall(r'style="[^"]*font-family[^"]*"', html_after_mso_clean)
print(f"\nAfter MSO clean - style attrs with font-family: {len(broken_check2)}")
for bc in broken_check2[:3]:
    if len(bc) > 200:
        print(f"  LONG MATCH ({len(bc)} chars): {bc[:100]}...{bc[-50:]}")
    else:
        print(f"  {bc}")

# Now compare with the full _sanitize_email_html
converter = EmailToPDFConverter()
full_sanitized = converter._sanitize_email_html(email.body_html)
broken_check3 = re.findall(r'style="[^"]*font-family[^"]*"', full_sanitized)
print(f"\nFull _sanitize_email_html - style attrs with font-family: {len(broken_check3)}")
for bc in broken_check3[:3]:
    if len(bc) > 200:
        print(f"  LONG MATCH ({len(bc)} chars): {bc[:100]}...{bc[-50:]}")
    else:
        print(f"  {bc}")

# Key question: are there any style attributes using SINGLE quotes in the raw HTML?
single_style = re.findall(r"style='[^']*'", raw)
print(f"\nSingle-quoted style attrs in RAW HTML: {len(single_style)}")
for ss in single_style[:3]:
    print(f"  {ss[:150]}")

# Are there Span tags with mixed case?
mixed_span = re.findall(r'<Span[^>]*>', raw)
print(f"\nMixed-case <Span> tags: {len(mixed_span)}")
for ms in mixed_span[:2]:
    print(f"  {ms[:150]}")

# Show the EXACT context around first font-family occurrence after fix
idx = html_after_ff_fix.find("font-family:'Aptos'")
if idx >= 0:
    start = max(0, idx - 50)
    end = min(len(html_after_ff_fix), idx + 150)
    print(f"\nContext around first fixed font-family:")
    print(f"  {html_after_ff_fix[start:end]!r}")
