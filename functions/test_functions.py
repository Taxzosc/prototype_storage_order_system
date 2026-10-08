import unittest
from openpyxl import load_workbook
from settings import FILE_LOCATION, EXCEL_FILE
from excel_functions import get_all_saved_unique_ordernumbers
from helper_func import convert_letter_to_index

class Testget_all_unique_ordernumbers(unittest.TestCase):
    wb = f"{FILE_LOCATION}/{EXCEL_FILE}"
    def test_correct_letter_to_index(self):
        index = convert_letter_to_index("g")
        self.assertEqual(index, 6)