# -*- coding: utf-8 -*-
import sys
import io
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

with open('curated_scholarships.json', 'r', encoding='utf-8') as f:
    items = json.load(f)

print(f"Total items before cleaning: {len(items)}")

# Identify invalid / note rows
cleaned = []
removed = []

for item in items:
    name = item.get("name", "").strip()
    name_ar = item.get("name_ar", "").strip()
    
    # Check if this row is a note, legend, or instruction
    is_note = False
    for bad_word in ["ملاحظات:", "اتأكدت منه", "اللينكات الأساسية", "(تقريبي) أو (متوقع)", "تاريخ إعداد الملف", "كيف تقرئين"]:
        if bad_word in name or bad_word in name_ar:
            is_note = True
            break
            
    if is_note or len(name) < 3 or not item.get("link"):
        # Let's check if link is present or if it's a real scholarship
        if is_note:
            removed.append(name)
            continue
            
    cleaned.append(item)

print(f"Removed {len(removed)} note rows:")
for r in removed:
    print(" -", r)

print(f"Total items after cleaning: {len(cleaned)}")

# Re-assign sequential IDs
for idx, item in enumerate(cleaned, 1):
    item["seq_no"] = idx

with open('curated_scholarships.json', 'w', encoding='utf-8') as f:
    json.dump(cleaned, f, ensure_ascii=False, indent=2)

print("Saved cleaned dataset to curated_scholarships.json")
