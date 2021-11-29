from excel_utils import *
from pythonutils.input_utils import *

user_input = True
workbook_source_path = ''

if user_input:
    print('')
    workbook_source_path = stripped_input('Enter/paste the path of the workbook you delete all of the Named Ranges from: ')
else:
    workbook_source_path = '/Users/georgekatsaros/Projects/Legendary/Config/Excel/y.Old Events/E273/CincoDeMayo/+E273-CD110+BlackoutBingo.xlsx'

delete_broken_named_ranges(workbook_source_path)

print('Done!')
