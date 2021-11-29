from excel_utils import *
import openpyxl
from pythonutils.input_utils import *

user_input = True
workbook_source_path = ''
workbook_target_path = ''
sheet_name_target = ''
source_row = 1
target_row = 1

if user_input:
    print('')
    workbook_source_path = stripped_input('Enter/paste the path of the workbook you want to transpose data FROM: ')
    workbook_target_path = stripped_input('Enter/paste the path of the workbook you want to transpose data TO: ')
    sheet_name_target = stripped_input('Enter/paste the name of the Sheet you want: ')
    source_row = stripped_input('Enter/paste the row you want to source the transposition from: ')
    target_row = stripped_input('Enter/paste the row you want to source the transposition to: ')
else:
    workbook_source_path = '/Users/georgekatsaros/Desktop/ExcelPlayground/PokemonAll.xlsx'
    workbook_target_path = '/Users/georgekatsaros/Desktop/ExcelPlayground/PokemonEvent.xlsx'
    sheet_name_target = 'Pokemon'
    source_row = 2
    target_row = 2

workbook_source = openpyxl.load_workbook(workbook_source_path, data_only=False)
workbook_target = openpyxl.load_workbook(workbook_target_path, data_only=False)

sheet_source = find_sheet_name_that_startswith(from_this_workbook=workbook_source, sheet_name=sheet_name_target)
sheet_target = find_sheet_name_that_startswith(from_this_workbook=workbook_target, sheet_name=sheet_name_target)

transpose_row(from_sheet=sheet_source, from_row=source_row, to_sheet=sheet_target, to_row=target_row)

workbook_target.save(workbook_target_path)
workbook_source.close()
workbook_target.close()

print('Done!')
