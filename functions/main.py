# this is the main workspace for this project
import datetime
from openpyxl import load_workbook
from settings import FILE_LOCATION, EXCEL_FILE
from excel_functions import check_product_number_in_database, get_all_saved_unique_ordernumbers, subtract_product_amount, add_to_bottom_history
from helper_func import convert_letter_to_index, check_if_already_exists

wb = load_workbook(filename = f"{FILE_LOCATION}/{EXCEL_FILE}")
sim_order = [2246101,[10001,"Bokashistrø, 1 kg",2]]
database = "Test Database"
historikk = "Test historikk"
unique_ordernumbers = get_all_saved_unique_ordernumbers(wb, historikk, "g")

#inside submit loop?
order_number_check = check_if_already_exists(unique_ordernumbers, sim_order[0])
if order_number_check == True:
    raise Exception("order number previously used")  #type of error and how its handled will be different in finished program with GUI
for i in range(1, len(sim_order)): #per product in order
    product_number = sim_order[i][0]
    amount = sim_order[i][2]
    new_amount = subtract_product_amount(wb,historikk, product_number, amount)
    #now i need to add to history bottom row: date, product number, product name, amount subtracted, new amount number, type of transaction, ordernumber
    add_to_bottom_history(wb,historikk,sim_order[0],sim_order[i],new_amount)
unique_ordernumbers.add(sim_order[0])
# add save xlsx line.