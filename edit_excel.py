from openpyxl import load_workbook
import sys

# This gets the file name from PowerShell
file_path = sys.argv[1]

# This gets the date from PowerShell
week_end = sys.argv[2]

# Open the Excel file
wb = load_workbook(file_path)

# Select first sheet
ws = wb.active

# Change cell B2
ws["B2"] = "WE " + week_end

# Save file
wb.save(file_path)

print("Excel updated successfully")
print(" Running Python")