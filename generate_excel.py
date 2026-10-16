# -*- coding: utf-8 -*-
import sys
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

with open('curated_scholarships.json', 'r', encoding='utf-8') as f:
    items = json.load(f)

print(f"Generating Excel file for {len(items)} scholarships...")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "دليل المنح والمنظمات"
ws.views.sheetView[0].rightToLeft = True

# Define styles
header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

category_colors = {
    "أفريقيا والدول النامية": "FEF3C7", # Light amber
    "المنح الحكومية والدولية الكبرى": "DBEAFE", # Light blue
    "مراكز ومعاهد التكنولوجيا الحيوية والعلوم الطبية": "DCFCE7", # Light green
    "منح الجامعات العالمية والمؤسسات المرموقة": "F3E8FF", # Light purple
    "محركات البحث وأدلة وظائف الدكتوراه": "FCE7F3", # Light pink
}

thin_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

headers = [
    "م",
    "اسم المنحة / المنظمة",
    "التصنيف والنطاق",
    "المرحلة الدراسية",
    "الرابط الرسمي",
    "الشروط الأساسية والمتطلبات",
    "ميعاد بداية التقديم",
    "آخر ميعاد للتقديم (Deadline)",
    "نوع التمويل",
    "ملاحظات هامة وإرشادات"
]

ws.append(headers)

# Format header row
for col_idx in range(1, len(headers) + 1):
    cell = ws.cell(row=1, column=col_idx)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

ws.row_dimensions[1].height = 32

for idx, item in enumerate(items, 1):
    row_data = [
        idx,
        item.get("name_ar") if item.get("name_ar") and item.get("name_ar") != item.get("name") else item.get("name"),
        item.get("category", "عام"),
        item.get("degree", "ماجستير ودكتوراه"),
        item.get("link", ""),
        item.get("requirements", ""),
        item.get("start_date", ""),
        item.get("deadline", ""),
        item.get("funding", "تمويل كامل أو جزئي"),
        item.get("notes", "")
    ]
    ws.append(row_data)
    row_num = idx + 1
    ws.row_dimensions[row_num].height = 42

    cat = item.get("category", "")
    bg_color = category_colors.get(cat, "FFFFFF")
    cat_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=row_num, column=col_idx)
        cell.font = Font(name="Calibri", size=10)
        cell.border = thin_border
        
        # Center align ID, category, degree, dates
        if col_idx in [1, 3, 4, 7, 8]:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        else:
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            
        # Highlight category cell
        if col_idx == 3:
            cell.fill = cat_fill
            cell.font = Font(name="Calibri", size=10, bold=True)

# Adjust column widths
col_widths = {
    1: 6,   # م
    2: 38,  # الاسم
    3: 25,  # التصنيف
    4: 18,  # الدرجة
    5: 35,  # الرابط
    6: 45,  # الشروط
    7: 22,  # البداية
    8: 26,  # الموعد النهائي
    9: 30,  # التمويل
    10: 40  # ملاحظات
}

for col_idx, width in col_widths.items():
    col_letter = get_column_letter(col_idx)
    ws.column_dimensions[col_letter].width = width

excel_path = "Scholarships_Master_PhD_Radwa_Complete.xlsx"
wb.save(excel_path)
print(f"Successfully generated {excel_path} with {len(items)} rows!")
