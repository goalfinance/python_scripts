import unittest
from excel.translator import * 
import numpy as np

class TestTranslator(unittest.TestCase):
    """Test translator.py"""
    def test_get_attendances_matrix(self):
        source_workbook = load_workbook("source_attendance.xlsx")
        attendances_matrix, members_full_name = get_attendances_matrix(source_workbook, 2019, 7)

        print(attendances_matrix)
        print(members_full_name)

if __name__ == "__main__":
    unittest.main()
