from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

filename = "dropdown_test.xlsx"

# Create workbook
wb = Workbook()
ws = wb.active

headers = ["Sender", "Date", "Time", "Wait Time", "Status"]
ws.append(headers)

# Create data validation
dv = DataValidation(
    type="list",
    formula1='"Not started,In Progress,Done"',
    allow_blank=True
)

# Add validation to sheet
ws.add_data_validation(dv)

# Add validation to E2:E100
dv.add("E2:E100")

# Append a test row
ws.append(["someone@example.com", "2025-07-01", "12:00:00", "0 days, 0 hours, 5 minutes", ""])

# Save workbook
wb.save(filename)
print(f"✅ Created {filename} with dropdown.")
