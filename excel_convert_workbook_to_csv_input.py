from pythonutils.input_utils import *
from pythonutils.yes_or_no_input import *
from excel_convert_workbook_to_csv import *
import os

user_input = True
convert_path_input = ''
recursive_input = True
input_export_path = ''

print("")
if user_input:
    convert_path_input = stripped_input('Enter/paste/drag and drop the workbook or directory you want to convert: ')
    recursive_input = False
    if os.path.isdir(convert_path_input):
        recursive_input = yes_or_no('Recursive?')

    input_export_path = stripped_input('Enter/paste/drag and drop the directory you want to export to: ')
else:
    convert_path_input = '/Users/georgekatsaros/Projects/Legendary/Config/Excel'
    recursive_input = True
    input_export_path = '/Users/georgekatsaros/Projects/LegendaryMirror/CSVs'

if os.path.isdir(convert_path_input):
    convert_dir(convert_path_input, input_export_path, recursive_input)
else:
    convert_workbook(convert_path_input, input_export_path)
