import sys, getopt
import time
import os
import openpyxl
from enum import Enum

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





if __name__ == "__main__":
    main()