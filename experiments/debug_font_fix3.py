"""Check for HTML entity encoded quotes in font-family."""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Read the raw RTF-extracted HTML
with open('/Users/omayo/Downloads/rfp5_test_pdfs/1_body.html', 'r') as f:
    raw = f.read()

# Check for &quot; in font-family
entity_ff = re.findall(r'font-family\s*:\s*&quot;[^&]*&quot;', raw)
print(f"font-family with &quot; entities: {len(entity_ff)}")
for ff in entity_ff[:3]:
    print(f"  {ff}")

# Check ALL occurrences of font-family with their surrounding context
all_ff = list(re.finditer(r'font-family\s*:', raw))
print(f"\nTotal font-family: occurrences: {len(all_ff)}")

# Categorize them
dquote = []
squote = []
entity = []
bare = []
for m in all_ff:
    pos = m.end()
    after = raw[pos:pos+50]
    if after.startswith('"'):
        dquote.append(after[:40])
    elif after.startswith("'"):
        squote.append(after[:40])
    elif after.startswith('&quot;'):
        entity.append(after[:40])
    else:
        bare.append(after[:40])

print(f"  Starting with double-quote: {len(dquote)}")
print(f"  Starting with single-quote: {len(squote)}")
print(f"  Starting with &quot;: {len(entity)}")
print(f"  Starting with bare word: {len(bare)}")

if dquote:
    print(f"\n  Double-quote examples:")
    for d in dquote[:3]:
        print(f"    {d!r}")
if entity:
    print(f"\n  Entity examples:")
    for e in entity[:3]:
        print(f"    {e!r}")
if bare:
    print(f"\n  Bare examples:")
    for b in bare[:3]:
        print(f"    {b!r}")

# Check if there's a pattern like style="font-size:11pt;font-family: (with more context)
# Look at a specific instance
idx = raw.find('font-family:')
if idx >= 0:
    start = max(0, idx - 30)
    end = min(len(raw), idx + 80)
    print(f"\nFirst occurrence context:")
    print(f"  {raw[start:end]!r}")
