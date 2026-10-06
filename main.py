# this is the main workspace for this project
from openpyxl import load_workbook

wb = load_workbook(filename = "files/test teoretisk database+historikk.xlsx")
sheets = wb.sheetnames
historikk = wb["historikk"]
database = wb["Database"]
print(database.max_column)
print(database.max_row)

all_prod_ids = []

for row in database:
    if type(row[0].value) != int:
        continue
    prod_id = row[0].value
    all_prod_ids.append(prod_id)
print(all_prod_ids)