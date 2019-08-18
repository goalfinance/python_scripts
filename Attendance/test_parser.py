import unittest
from excel.parser import * 
import numpy as np

class TestParser(unittest.TestCase):
    """Test parser.py"""
    def test_parse_source_file(self):
        source_workbook = load_workbook("source_attendance.xlsx")
        attendances_matrix = get_attendances_matrix(source_workbook, 2019, 7)

        print(attendances_matrix)

if __name__ == "__main__":
    unittest.main()
