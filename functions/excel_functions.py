from openpyxl import load_workbook, Workbook
from helper_func import convert_letter_to_index
from settings import FILE_LOCATION, EXCEL_FILE
import datetime



def get_all_saved_unique_ordernumbers(workbook: Workbook, history_sheet_name: str, column: str | int) -> set:
    """Returns a set of order numbers in this sheet

    Loops through specified column in sheet to find all ordernumbers
    then placing them in a set to avoid duplicates
    if parameter column is a letter, it will be converted to the corresponding number,
    assuming alphabet limit of a-z(26 char)
    index 0-25
    """
    sheet = workbook[history_sheet_name]

    if type(column) == str:
        column = convert_letter_to_index(column)

    order_numbers = set()
    for row in sheet:
        if type(row[column].value) == float or type(row[column].value) ==int:
            order_number = row[column].value
            order_numbers.add(order_number)
    return order_numbers


def check_product_number_in_database(workbook: Workbook, database_sheet_name: str, column: str | int, product_number: int|float) -> bool:
    database = workbook[database_sheet_name]

    if type(column) == str:
        column = convert_letter_to_index(column)

    for row in database:
        if type(row[column].value) == float or type(row[column].value) ==int:
            if row[column].value == product_number:
                return True
    return False
def find_product_amount(workbook: Workbook, database_sheet_name: str, product_number: int|float)-> int|float:
    """Returns the current amount for a spesific product
    int or float
    assumes first row is always titles
    """
    database = workbook[database_sheet_name]

    for i in range(1,database.max_row):
        if type(database.cell(row=i, column=1).value) == float or type(database.cell(row=i, column=1).value) == int:
            if database.cell(row=i, column=1).value == product_number:
                return database.cell(row=i, column=3).value
    # TODO needs an error for when it fails to find product. invalid product

def subtract_product_amount(workbook: Workbook, database_sheet_name: str, product_number: int|float, subtractable_amount = int)-> int|float:
    """Subtracts amount from one product in worksheet, and Returns the new value
    loops through product database until product is found,
    then subtracts the subtractable amount from the current amount.
    assumes first row is always titles
    
    """
    database = workbook[database_sheet_name]

    for i in range(1,database.max_row):
        if type(database.cell(row=i, column=1).value) == float or type(database.cell(row=i, column=1).value) == int:
            if database.cell(row=i, column=1).value == product_number:
                # product_cell_value = database.cell(row=i, column=3).value #does nothing?
                new_value = database.cell(row=i, column=3).value - subtractable_amount
                database.cell(row=i, column=3).value = new_value
                return new_value
    # TODO needs an error for when it fails to find product. invalid product


def add_to_bottom_history(workbook: Workbook, history_sheet_name: str, order_number: int|float, order_product: list, new_amount: int|float):
    """adds a history of a product movement. 
    limited to "Ordre nettbutikk" for now
    
    """#TODO make  this work with other types of movement. (edit order, warehouse delivery, returns etc)
    #TODO change how bottom row works. make a function that finds bottom safe row. use "history_grab_bottom_row"
    history = workbook[history_sheet_name]
    bottom_empty_row = history.max_row+1
    time = datetime.datetime.now()
    prod_number = order_product[0]
    prod_name = order_product[1]
    amount_subtracted = -order_product[2]
    output_datetime = time.strftime("%d.%m.%Y,%H:%M:%S")

    #A:date, B:product number, C:product name, D:amount subtracted, E:new amount number, F:type of transaction, G:ordernumber
    history[f"A{bottom_empty_row}"] = output_datetime
    history[f"B{bottom_empty_row}"] = prod_number
    history[f"C{bottom_empty_row}"] = prod_name
    history[f"D{bottom_empty_row}"] = amount_subtracted
    history[f"E{bottom_empty_row}"] = new_amount
    history[f"F{bottom_empty_row}"] = "Ordre nettbutikk"
    history[f"G{bottom_empty_row}"] = order_number
    #TODO add worksheet save or not. add the save for code in main under the "submit" loop
    

def history_grab_bottom_or_spesific_row(workbook: Workbook, history_sheet_name: str, row: None | int = None)-> list:
    """Returns the values of the bottom or the spesified row, between a and g
    defaults to bottom unless max_row is spesified
    """
    #made this to check if other functions works, can now change it to be able to find bottom usable row.
    history = workbook[history_sheet_name]
    bottom_empty_row = history.max_row
    if row:
        bottom_empty_row = row
    return [history[f"A{bottom_empty_row}"].value,history[f"B{bottom_empty_row}"].value,history[f"C{bottom_empty_row}"].value,history[f"D{bottom_empty_row}"].value,history[f"E{bottom_empty_row}"].value,history[f"F{bottom_empty_row}"].value,history[f"G{bottom_empty_row}"].value]

def find_bottom_usable_row(workbook: Workbook, sheet_name: str, row_number: int = 0)-> int:
    """Returns the int of the bottom usable row.
    grabs the inital bottom row using sheet.max_row, checks if its properly empty between col a-col g,
    if not, it will go to a previous level until it finds a row with anything containing it.

    if any cell in the grabbed row contains a value, it will return that row number+1.
    
    this means if all but 1 cell has no value, it will still return this as the bottom row number
    """
    sheet = workbook[sheet_name]
    maxrow = sheet.max_row
    if maxrow-row_number <= 0: #just incase someone inputs a bad number from the initial call.
        return 1
    bottom_row = history_grab_bottom_or_spesific_row(workbook, sheet_name, maxrow+row_number)
    print(f"maxrow should be{maxrow-row_number}")
    print(bottom_row)
    for cell in bottom_row:
        if cell != None: 
            print("found nonempty cell")
            return (maxrow+row_number)+1
    print(f"found no nonempty row going down to level {row_number}")
    return find_bottom_usable_row(workbook, sheet_name, row_number-1)