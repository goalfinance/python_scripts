import openpyxl
import os

def load_workbook(file_path):
    if file_path == None:
        return None
    if (os.path.exists(file_path)):
        return openpyxl.load_workbook(file_path)
    else:
        return None

