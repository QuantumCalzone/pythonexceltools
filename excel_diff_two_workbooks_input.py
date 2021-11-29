from excel_diff_two_workbooks import *
from pythonutils.input_utils import *
from pythonutils.yes_or_no_input import *

user_input = True
execute_conversion = True


if user_input:
    print('')
    workbook_path_1_input = stripped_input('Enter/paste/drag and drop the FIRST Excel Workbook file you want to compare: ')
    workbook_path_2_input = stripped_input('Enter/paste/drag and drop the SECOND Excel Workbook file you want to compare: ')
    csv_diff_output_input = stripped_input('Enter/paste/drag and drop the directory you want to export the diff to: ')
    diff_formulas_input = yes_or_no('Show formulas?')
    only_show_diff_input = yes_or_no("Only show diff?")
else:
    workbook_path_1_input = '/Users/georgekatsaros/Downloads/Poke.xlsx'
    workbook_path_2_input = '/Users/georgekatsaros/Downloads/Poke 2.xlsx'
    csv_diff_output_input = '/Users/georgekatsaros/Downloads/'
    only_show_diff_input = True


while execute_conversion:
    workbook_diff(workbook_path_1_input, workbook_path_2_input, csv_diff_output_input, only_show_diff_input)
    execute_conversion = yes_or_no('Convert Again?')
