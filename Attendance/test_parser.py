import unittest
from excel.parser import *

class TestParser(unittest.TestCase):
    """Test parser.py"""
    def test_parse_source_file(self):
        source_workbook = load_workbook(".source_attendance.xlsx")
        if source_workbook != None and does_monsheet_exist(source_workbook, 7):
            pass


if __name__ == "__main__":
    unittest.main()
