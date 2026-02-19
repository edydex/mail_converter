"""Debug script to verify font-family fix is working in final HTML."""
import sys, os, re, logging
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
logging.basicConfig(level=logging.ERROR)

from core.eml_parser import EMLParser
from core.email_to_pdf import EmailToPDFConverter

parser = EMLParser()
email = parser.parse_file('/Users/omayo/Downloads/rfp5_test_emls/1.eml')

converter = EmailToPDFConverter()

# Build the FULL document that gets sent to WeasyPrint
full_html = converter._build_html_document(email, include_headers=True)
with open('/Users/omayo/Downloads/rfp5_test_pdfs/1_full.html', 'w') as f:
    f.write(full_html)

# Check for double-quoted font-family in the full document
remaining = re.findall(r'font-family\s*:\s*"[^"]*"', full_html)
print(f"Double-quoted font-family in full doc: {len(remaining)}")
for r in remaining[:5]:
    print(f"  {r}")

# Also check for the broken pattern in inline styles
broken_styles = re.findall(r'style="[^"]*font-family[^"]*"', full_html)
print(f"\nStyle attrs containing font-family: {len(broken_styles)}")
for s in broken_styles[:3]:
    print(f"  {s[:100]}")

# Check body-level style blocks
body_styles = re.findall(r'<style[^>]*>(.*?)</style>', full_html, re.DOTALL)
print(f"\nStyle blocks in full doc: {len(body_styles)}")
for i, block in enumerate(body_styles):
    dq_ff = re.findall(r'font-family\s*:\s*"[^"]*"', block)
    print(f"  Block {i}: {len(dq_ff)} double-quoted font-family")
    if dq_ff:
        for ff in dq_ff[:2]:
            print(f"    {ff}")

# Check for the specific broken pattern where double quote breaks the style attr
# Pattern: style="...font-family:"FontName"..."  (inner double quotes)
broken_ff = re.findall(r'style="[^"]*?font-family\s*:\s*"', full_html)
print(f"\nBroken style attrs (font-family terminates style): {len(broken_ff)}")
for b in broken_ff[:3]:
    print(f"  {b}")
