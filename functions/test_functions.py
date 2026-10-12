import unittest
from openpyxl import load_workbook
from settings import FILE_LOCATION, EXCEL_FILE
from excel_functions import get_all_saved_unique_ordernumbers
from helper_func import convert_letter_to_index, check_if_already_exists

class Test_get_all_unique_ordernumbers(unittest.TestCase):
    
    def test_correct_letter_to_index(self):
        index = convert_letter_to_index("g")
        self.assertEqual(index, 6)

    def test_more_letters_to_index(self):
        index = convert_letter_to_index("p")
        self.assertEqual(index, 15)

    def test_get_all_unique_ordernumbers(self):
        wb = load_workbook(filename=f"{FILE_LOCATION}/{EXCEL_FILE}")
        extracted_order_nums = get_all_saved_unique_ordernumbers(wb, "Test historikk", "g")
        orders = {2246033,2246032,2246009,2245984,2241747}
        self.assertEqual(extracted_order_nums,orders)

class test_check_if_order_number_is_used(unittest.TestCase):

    def test_ord_num_exists_short_list(self):
        wb = load_workbook(filename=f"{FILE_LOCATION}/{EXCEL_FILE}")
        extracted_order_nums = get_all_saved_unique_ordernumbers(wb, "Test historikk", "g")
        exists = check_if_already_exists(extracted_order_nums, 2246033)
        self.assertEqual(exists, True)

    def test_not_ord_num_exists_short_list(self):
        wb = load_workbook(filename=f"{FILE_LOCATION}/{EXCEL_FILE}")
        extracted_order_nums = get_all_saved_unique_ordernumbers(wb, "Test historikk", "g")
        exists = check_if_already_exists(extracted_order_nums, 2246030)
        self.assertEqual(exists, False)
        
