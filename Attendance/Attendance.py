import sys, getopt
import time
import os
import openpyxl
from enum import Enum
import re

class Command(Enum):
    addMember = 1
    initial = 2
    recordLeaving = 3

class _const:
    class ConstError(TypeError):pass
    def __setattr__(self, name, value):
        if name in self.__dict__:
            raise self.ConstError("Can't rebind const (%s)" %name)
        self.__dict__[name] = value
class MissingParamsError(TypeError):
    pass

class ExcelFileOperationError(TypeError):
    pass

const = _const()
const.APP_PARAMS_MEMBER_NAME = 'member_name'
const.APP_PARAMS_LEAVING_DATE = 'leaving_date'
const.APP_PARAMS_FILE_NAME = 'file_name'
const.EXCEL_WORKBOOK_MEMBER_SHEET_NAME = 'member list'

app_params = dict(member_name="", leaving_date="", file_name="")

def month_to_str(month):
    month_str = str(month)
    if len(month_str) < 2:
        month_str = '0' + month_str 
    return month_str

#Check if the sheet for the month exists
def does_monsheet_exist(workbook, mon):
    month_str = month_to_str(mon)
    sheet_names = workbook.get_sheet_names()
    for sheet_name in sheet_names:
        if sheet_name == month_str:
            return True
    
    return False

def get_month_sheet(workbook, mon):
    month_str = month_to_str(mon)
    if does_monsheet_exist(workbook, mon) == True:
        return workbook.get_sheet_by_name(month_str)
    else:
        return None

def get_current_month_num():
    localtime = time.localtime(time.time())
    return localtime.tm_mon

def get_previous_month_num():
    current_month_num = get_current_month_num()
    previous_month_num = current_month_num - 1
    if previous_month_num < 1 :
        previous_month_num = 12
    return previous_month_num

def perform_add_member(workbook, app_params):
    member_name = app_params[const.APP_PARAMS_MEMBER_NAME]
    if member_name == None or member_name == '':
        raise MissingParamsError("The name of member you want to add is missing, please assigning it by using '-m', '--member_name'")
    try:
        member_list_sheet = workbook.get_sheet_by_name(const.EXCEL_WORKBOOK_MEMBER_SHEET_NAME)
    except KeyError:
        member_list_sheet = workbook.create_sheet(const.EXCEL_WORKBOOK_MEMBER_SHEET_NAME, index=0)
    
    new_row = [member_name]
    member_list_sheet.append(new_row)
    workbook.save(app_params[const.APP_PARAMS_FILE_NAME])

def does_member_exist(workbook, member_name):
    try:
        member_list_sheet = workbook.get_sheet_by_name(const.EXCEL_WORKBOOK_MEMBER_SHEET_NAME)
        for row in member_list_sheet.rows:
            if re.search(member_name, row[0].value, re.IGNORECASE) != None:
                return True
        return True
    except KeyError:
        return False


def perform_initial_attendance(workbook, member_name, month):
    month_sheet = get_month_sheet(workbook, month)
    if month_sheet == None:
        month_sheet = workbook.create_sheet(month_to_str(month))
    
    if does_member_exist(workbook, member_name) == False:
        print("The member whose name is '" + member_name + "' doesn't exist in the member list, you can add member by option '-a'")
        return
    
    for row in month_sheet.rows:
        if re.search(member_name, row[0].value, re.IGNORECASE) == None:
            pass

    workbook.save(app_params[const.APP_PARAMS_FILE_NAME])





def main():
    print(sys.argv[0])
    try:
        opts, args = getopt.getopt(sys.argv[1:], "airm:l:f:", ["member_name=", "leaving_date=", "file"])
    except getopt.GetoptError as err:
        print(err)
        sys.exit(2)
    
    command = Command.initial
    for o, a in opts:
        if o in ("-m", "--member_name"):
            app_params[const.APP_PARAMS_MEMBER_NAME] = a
        elif o in ("-l", "--leaving_date"):
            app_params[const.APP_PARAMS_LEAVING_DATE] = a
        elif o in ("-f", "--file"):
            app_params[const.APP_PARAMS_FILE_NAME] = a
        elif o in ("-a"):
            command = Command.addMember
        elif o in ("-i"):
            command = Command.initial
        elif o in ("-r"):
            command = Command.recordLeaving
    
    
    if app_params[const.APP_PARAMS_LEAVING_DATE] != None and app_params[const.APP_PARAMS_LEAVING_DATE] != "":
        try:
            leaving_date = time.strptime(app_params[const.APP_PARAMS_LEAVING_DATE], "%Y-%m-%d")
        except ValueError as err:
            print("The format of leaving_date is incorrect[" + str(err) + "]")
            sys.exit(2)    
        
    print("member_name = [" + app_params[const.APP_PARAMS_MEMBER_NAME] + "]")
    print("leaving_date = [" + app_params[const.APP_PARAMS_LEAVING_DATE] + "]")
    print("file_name = [" + app_params[const.APP_PARAMS_FILE_NAME] + "]")

    file_name = app_params[const.APP_PARAMS_FILE_NAME]
    if file_name == None or file_name == '':
        print("The attentance file can't not be empty, assigning it by using '-f' or '--file-name'")
        sys.exit(2)
    
    if os.path.exists(file_name):
        book = openpyxl.load_workbook(file_name)
    else:
        book = openpyxl.Workbook(write_only=True)
        book.save(file_name)
    
    if command == Command.addMember:
        perform_add_member(book, app_params)
    elif command == Command.initial:
        perform_initial_attendance(book, app_params)





if __name__ == "__main__":
    main()