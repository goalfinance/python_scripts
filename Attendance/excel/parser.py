import openpyxl
import os
from goalfinance.utils.calendar import month_to_str
import numpy as np
import re

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
    source_attendances = dict()
    for row in month_sheet.rows:
        member_name = row[1].value
        members_attendances_records = []
        if member_name not in source_attendances:
            source_attendances[member_name] = []
            members_attendances_records = source_attendances[member_name]
        else:
            members_attendances_records = source_attendances[member_name]
            if members_attendances_records == None:
                members_attendances_records = []
                source_attendances[member_name] = members_attendances_records

        members_attendances_records.append(row)
    
    attendances_matrix = dict()

    for member_name in list(source_attendances):
        row_cnt = 0
        column_cnt = month_sheet.max_column
        if source_attendances[member_name] != None:
            row_cnt = len(source_attendances[member_name])
            if row_cnt > 0:
                matrix = np.zeros((row_cnt, column_cnt - 2))
                attendances_matrix[member_name] = matrix
                x = 0
                for row in source_attendances[member_name]:
                    y = 0
                    for column_number in range(2, column_cnt - 1):
                        if row[column_number].value != None and re.search("absen*", row[column_number].value, re.IGNORECASE) != None:
                            matrix[x, y] = 1
                        y += 1
                    x += 1
    
    return attendances_matrix


           

