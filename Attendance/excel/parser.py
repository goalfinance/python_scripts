import openpyxl
import os
from goalfinance.utils.calendar import month_to_str

def load_workbook(file_path):
    if file_path == None:
        return None
    if (os.path.exists(file_path)):
        return openpyxl.load_workbook(file_path)
    else:
        return None

#Check if the sheet for the month exists
def does_monsheet_exist(workbook, mon):
    month_str = month_to_str(mon)
    sheet_names = workbook.sheetnames
    for sheet_name in sheet_names:
        if sheet_name == month_str:
            return True
    
    return False

def get_month_sheet(workbook, mon):
    month_str = month_to_str(mon)
    if does_monsheet_exist(workbook, mon) == True:
        return workbook[month_str]
    else:
        return None

def source_attendance_group_by_member(source_workbook, month):
    month_sheet = get_month_sheet(source_workbook, month)
    if month_sheet == None:
        return None
    source_attendances = []
    for row in month_sheet.rows:
        

