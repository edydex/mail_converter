"""Debug script to verify font-family fix is working."""
import sys, os, re, logging
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
logging.basicConfig(level=logging.WARNING)

from core.eml_parser import EMLParser
from core.email_to_pdf import EmailToPDFConverter

parser = EMLParser()
email = parser.parse_file('/Users/omayo/Downloads/rfp5_test_emls/1.eml')

converter = EmailToPDFConverter()
sanitized = converter._sanitize_email_html(email.body_html)

# Save sanitized HTML for inspection
with open('/Users/omayo/Downloads/rfp5_test_pdfs/1_sanitized.html', 'w') as f:
    f.write(sanitized)

# Check style blocks
style_blocks = re.findall(r'<style[^>]*>(.*?)</style>', sanitized, re.DOTALL | re.IGNORECASE)
print(f"Style blocks: {len(style_blocks)}")
for i, block in enumerate(style_blocks):
    print(f"\n--- Style block {i} (first 500 chars) ---")
    print(block[:500])

# Check for remaining problem patterns
remaining_dq_ff = re.findall(r'font-family\s*:\s*"[^"]*"', sanitized)
print(f"\nRemaining double-quoted font-family in HTML body: {len(remaining_dq_ff)}")
for ff in remaining_dq_ff[:3]:
    print(f"  {ff}")

# Check font-family in inline styles (should use single quotes)
inline_ff = re.findall(r"font-family\s*:[^;}{]+", sanitized)
print(f"\nAll font-family values ({len(inline_ff)} total), first 10:")
for ff in inline_ff[:10]:
    print(f"  {ff.strip()}")

# Now also check what the full HTML document looks like
full_html = converter._build_html_document(email, include_headers=True)
with open('/Users/omayo/Downloads/rfp5_test_pdfs/1_full.html', 'w') as f:
    f.write(full_html)

# Check if full doc has the problem
full_dq_ff = re.findall(r'font-family\s*:\s*"[^"]*"', full_html)
print(f"\nIn full document: remaining double-quoted font-family: {len(full_dq_ff)}")
