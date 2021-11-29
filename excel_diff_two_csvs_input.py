from excel_diff_two_csvs import *
from pythonutils.input_utils import *
from pythonutils.yes_or_no_input import *

user_input = True

if user_input:
    print('')
    csv_path_1_input = stripped_input('Enter/paste/drag and drop the FIRST .csv file you want to compare: ')
    csv_path_2_input = stripped_input('Enter/paste/drag and drop the SECOND .csv file you want to compare: ')
    csv_diff_output_input = stripped_input('Enter/paste/drag and drop the directory you want to export the diff to: ')
    diff_formulas_input = yes_or_no('Show formulas?')
    only_show_diff_input = yes_or_no('Only show diff?')
else:
    csv_path_1_input = '/Users/georgekatsaros/Desktop/Exports/Book1/Sheet1.csv'
    csv_path_2_input = '/Users/georgekatsaros/Desktop/Exports/Book1/Sheet4.csv'
    csv_diff_output_input = '/Users/georgekatsaros/Desktop/'
    only_show_diff_input = True

csv_diff(csv_path_1_input, csv_path_2_input, csv_diff_output_input, only_show_diff_input)
