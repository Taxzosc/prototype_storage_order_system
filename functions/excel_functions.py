from openpyxl import load_workbook, Workbook
from helper_func import convert_letter_to_index
from settings import FILE_LOCATION, EXCEL_FILE



def get_all_saved_unique_ordernumbers(workbook: Workbook, sheetname: str, column: str | int) -> set:
    """Returns a set of order numbers in this sheet

    Loops through specified column in sheet to find all ordernumbers
    then placing them in a set to avoid duplicates
    if parameter column is a letter, it will be converted to the corresponding number,
    assuming alphabet limit of a-z(26 char)
    index 0-25
    """
    sheet = workbook[sheetname]

    if type(column) == str:
        column = convert_letter_to_index(column)

    order_numbers = set()
    for row in sheet:
        if type(row[column].value) != float:
            continue
        order_number = row[column].value
        order_numbers.add(order_number)
    return order_numbers