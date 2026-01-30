"""
Debug script to compare property tags between working and failing scripts.
"""

from win32com.mapi import mapitags

print("Property tag comparison:")
print("=" * 60)

# Working script uses these from mapitags module:
tags_to_check = [
    ("PR_SUBJECT_A", "mapitags.PR_SUBJECT_A"),
    ("PR_BODY_A", "mapitags.PR_BODY_A"),
    ("PR_MESSAGE_CLASS_A", "mapitags.PR_MESSAGE_CLASS_A"),
    ("PR_MESSAGE_FLAGS", "mapitags.PR_MESSAGE_FLAGS"),
    ("PR_MESSAGE_DELIVERY_TIME", "mapitags.PR_MESSAGE_DELIVERY_TIME"),
    ("PR_CLIENT_SUBMIT_TIME", "mapitags.PR_CLIENT_SUBMIT_TIME"),
    ("PR_SENDER_NAME_A", "mapitags.PR_SENDER_NAME_A"),
    ("PR_SENDER_EMAIL_ADDRESS_A", "mapitags.PR_SENDER_EMAIL_ADDRESS_A"),
]

# My hardcoded values:
my_values = {
    "PR_SUBJECT_A": 0x0037001E,
    "PR_BODY_A": 0x1000001E,
    "PR_MESSAGE_CLASS_A": 0x001A001E,
    "PR_MESSAGE_FLAGS": 0x0E070003,
    "PR_MESSAGE_DELIVERY_TIME": 0x0E060040,
    "PR_CLIENT_SUBMIT_TIME": 0x00390040,
    "PR_SENDER_NAME_A": 0x0C1A001E,
    "PR_SENDER_EMAIL_ADDRESS_A": 0x0C1F001E,
}

for name, attr in tags_to_check:
    try:
        module_val = getattr(mapitags, name)
        my_val = my_values.get(name, "N/A")
        match = "✓" if module_val == my_val else "❌ MISMATCH!"
        print(f"{name}:")
        print(f"  mapitags: {hex(module_val)} ({module_val})")
        print(f"  mine:     {hex(my_val)} ({my_val})")
        print(f"  {match}")
    except AttributeError:
        print(f"{name}: NOT FOUND in mapitags")
    print()

print("\n" + "=" * 60)
print("Now let's look at all PR_* that end with _A:")
print("=" * 60)

for name in dir(mapitags):
    if name.startswith("PR_") and name.endswith("_A"):
        val = getattr(mapitags, name)
        print(f"{name} = {hex(val)}")
