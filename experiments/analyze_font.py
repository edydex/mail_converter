"""Analyze RTF HTML font-family issue."""
import re, sys

html_file = sys.argv[1] if len(sys.argv) > 1 else '/Users/omayo/Downloads/rfp5_test_pdfs/1_body.html'

with open(html_file) as f:
    html = f.read()

# Find all font-family usage in inline styles
matches = re.findall(r'font-family:\s*([^;"]{0,80})', html)
empty = [m for m in matches if not m.strip()]
non_empty = [m for m in matches if m.strip()]
print(f'Empty font-family: {len(empty)}')
print(f'Non-empty font-family: {len(non_empty)}')
print()
print('Non-empty values:')
for v in sorted(set(non_empty)):
    print(f'  [{v}]')

# Check context around empty font-family declarations
pattern = re.compile(r'font-family:\s*[;"]')
for i, m in enumerate(pattern.finditer(html)):
    start = max(0, m.start() - 30)
    end = min(len(html), m.end() + 30)
    context = html[start:end].replace('\n', ' ')
    if i < 5:
        print(f'\nEmpty font-family context #{i+1}: ...{context}...')

# Check what the RTF raw bytes look like
# Look for specific patterns around quotes
print(f'\n--- Quote analysis ---')
print(f"Single quote (apostrophe) count: {html.count(chr(39))}")
print(f"Left single quote count: {html.count(chr(8216))}")
print(f"Right single quote count: {html.count(chr(8217))}")
print(f"&quot; count: {html.count('&quot;')}")

# Find what's between font-family: and the next ; 
# Look for cases where a quote is immediately followed by a font name
font_ctx = re.findall(r'font-family:\s*(.{0,3})', html)
print(f'\nFirst chars after font-family:')
for fc in sorted(set(font_ctx)):
    count = font_ctx.count(fc)
    print(f'  [{repr(fc)}] x{count}')
