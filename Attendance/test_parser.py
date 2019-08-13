import unittest
from excel.parser import * 

class TestParser(unittest.TestCase):
    """Test parser.py"""
    def test_parse_source_file(self):
        source_workbook = load_workbook("source_attendance.xlsx")
        attendance_group_by_member = get_attendance_matrix(source_workbook, 2019, 7)

        print(attendance_group_by_member)
        


if __name__ == "__main__":
    unittest.main()
