import sys, getopt
import time
import os
import openpyxl
from enum import Enum


class Command(Enum):
    addMember = 1
    initial = 2
    recordLeaving = 3

def main():
    print(sys.argv[0])
    try:
        opts, args = getopt.getopt(sys.argv[1:], "airm:l:f:", ["member_name=", "leaving_date=", "file"])
    except getopt.GetoptError as err:
        print(err)
        sys.exit(2)
    app_params = dict(member_name="", leaving_date="", file_name="")
    command = Command.initial
    for o, a in opts:
        if o in ("-m", "--member_name"):
            app_params["member_name"] = a
        elif o in ("-l", "--leaving_date"):
            app_params["leaving_date"] = a
        elif o in ("-f", "--file"):
            app_params["file_name"] = a
        elif o in ("-a"):
            command = Command.addMember
        elif o in ("-i"):
            command = Command.initial
        elif o in ("-r"):
            command = Command.recordLeaving
    
    

    try:
        leaving_date = time.strptime(app_params["leaving_date"], "%Y-%m-%d")
    except ValueError as err:
        print("The format of leaving_date is incorrect[" + str(err) + "]")
        sys.exit(2)    
        
    print("member_name = [" + app_params["member_name"] + "]")
    print("leaving_date = [" + app_params["leaving_date"] + "]")

    file_name = app_params["file_name"]
    print("file_name = [" + app_params["file_name"] + "]")
    if os.path.exists(file_name):
        book = openpyxl.load_workbook(file_name)
    else:
        book = openpyxl.Workbook(write_only=True)
        book.save(file_name)
    
    
    book.save(file_name)



if __name__ == "__main__":
    main()