# this is the main workspace for this project
from openpyxl import load_workbook

wb = load_workbook(filename = "files/test teoretisk database+historikk.xlsx")
sheets = wb.sheetnames
historikk = wb["historikk"]
database = wb["Database"]
print(database.max_column)
print(database.max_row)