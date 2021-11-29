from pythonutils.input_utils import *
from excel_mirror import *

user_input = True
if user_input:
    print("")
    excel_dir_path_input = stripped_input("Enter/paste/drag and drop the directory of the excel files you want to mirror: ")
    mirror_dir_path_input = stripped_input("Enter/paste/drag and drop the directory of the mirror: ")
else:
    excel_dir_path_input = "/Users/georgekatsaros/Desktop/ExcelWorkbooks"
    mirror_dir_path_input = "/Users/georgekatsaros/Desktop/ExcelWorkbooksMirror"

mirror_workbooks(excel_dir_path_input, mirror_dir_path_input)
