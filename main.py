import openpyxl
import re

# Sample draft marks (manually copied from GCR)
gcr_marks = {
    "Huzaifa Nasir": "88/100",
    "231-0079 Awwab Ahmad": "77/100",
    "Abdullah Tariq": "85/100",
    "Ahmar Ali": "90/100",
    "Amna Shahid": "80/100",
    "Faisal Yaseen": "74/100",
    "Faateh Syed": "79/100",
    "Gri Raj": "92/100"
    # Add more...
}

# Preprocess GCR: extract last 4 digits of roll numbers
processed_gcr = {}
for name, mark in gcr_marks.items():
    match = re.search(r'(\d{4})', name)
    key = match.group(1) if match else name.strip()
    processed_gcr[key] = mark

# Load Excel
wb = openpyxl.load_workbook("students.xlsx")
ws = wb.active

# Add column header if needed
header_row = [cell.value for cell in ws[1]]
if "HW3 Marks" not in header_row:
    ws.cell(row=1, column=len(header_row)+1).value = "HW3 Marks"
    mark_col = len(header_row) + 1
else:
    mark_col = header_row.index("HW3 Marks") + 1

# Match and insert marks
for row in ws.iter_rows(min_row=2, values_only=False):
    roll_cell = row[1]  # Assuming Roll Number is column B
    name_cell = row[0]  # Assuming Name is column A

    roll_match = re.search(r'(\d{4})', str(roll_cell.value))
    roll_suffix = roll_match.group(1) if roll_match else None

    mark = None
    if roll_suffix and roll_suffix in processed_gcr:
        mark = processed_gcr[roll_suffix]
    elif name_cell.value and name_cell.value.strip() in processed_gcr:
        mark = processed_gcr[name_cell.value.strip()]

    if mark:
        row[mark_col - 1].value = mark

# Save Excel
wb.save("students_updated.xlsx")
print("✅ Marks added successfully.")
