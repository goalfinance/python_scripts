import unittest
from excel.parser import *

class TestParser(unittest.TestCase):
    """Test parser.py"""
    def test_parse_source_file(self):
        source_workbook = load_workbook(".source_attendance.xlsx")
        attendance_group_by_member = source_attendance_group_by_member(source_workbook, 7)
        


if __name__ == "__main__":
    unittest.main()
