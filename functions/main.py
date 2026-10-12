# this is the main workspace for this project
from openpyxl import load_workbook
from settings import FILE_LOCATION, EXCEL_FILE

wb = load_workbook(filename = f"{FILE_LOCATION}/{EXCEL_FILE}")
print(f"{FILE_LOCATION}/{EXCEL_FILE}")