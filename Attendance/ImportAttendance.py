from getopt import getopt
from getopt import GetoptError
import sys
from goalfinance.utils.utils import Const
from excel.parser import load_workbook


app_params = dict(source="", target="")
const = Const()
const.APP_PARAMS_SOURCE_FILE = "source"
const.APP_PARAMS_TARGET_FILE = "target"
const.APP_PARAMS_MONTH = "month"

def main():
    try:
        opts, args = getopt(sys.argv[1:], "s:t:m:", [const.APP_PARAMS_SOURCE_FILE, const.APP_PARAMS_TARGET_FILE, const.APP_PARAMS_MONTH])
    except GetoptError as opt_error:
        print(opt_error)
        exit(2)

    for o, a in opts:
        if o in ("-s", "--source"):
            app_params[const.APP_PARAMS_SOURCE_FILE] = a
        elif o in ("-t", "--target"):
            app_params[const.APP_PARAMS_TARGET_FILE] = a
        elif o in ("-m", "-month"):
            app_params[const.APP_PARAMS_MONTH] = a
    
    if load_workbook(app_params[const.APP_PARAMS_SOURCE_FILE]) == None:
        print("The source file should not be absent, please assigning it by using '-s' or '--source'")
        sys.exit(2)

    if load_workbook(app_params[const.APP_PARAMS_TARGT_FILE]) == None:
        print("The target file should not be absent, please assigning it by using '-t' or '--target'")
        sys.exit(2)
    if app_params[const.APP_PARAMS_MONTH] == None:
        print("The month of attendance should not be absent, please assigning it by using '-m' or '--month'")


    

    

     




