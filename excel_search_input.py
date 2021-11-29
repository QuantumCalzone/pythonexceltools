from excel_search import *
from pythonutils.input_utils import *
from pythonutils.yes_or_no_input import *

print("")

input_path = stripped_input("Enter/paste/drag and drop the destination you wish to search: ")
if input_path == "":
    input_path = "/Users/georgekatsaros/Projects/Legendary/Config/Excel"
input_term = stripped_input("Enter/paste the term you wish to search: ")
input_search_formula = yes_or_no("Search formula?")
input_case_sensitive = yes_or_no("Case sensitive?")
input_match_entire_val = yes_or_no("Match entire value?")
input_sheet_name_contains = stripped_input("Optional | Enter/paste a name searched -SHEETS MUST- have: ")
input_sheet_name_excludes = stripped_input("Optional | Enter/paste a name searched -SHEETS must NOT- have: ")

results = []

if os.path.isdir(input_path):
    input_workbook_name_contains = stripped_input("Optional | Enter/paste a name searched -WORKBOOKS MUST- have: ")
    input_workbook_name_excludes = stripped_input("Optional | Enter/paste a name searched -WORKBOOKS must NOT- "
                                                  "have: ")

    input_recursive = yes_or_no("Search recursively?")

    print("")
    results = search_dir(path=input_path, term=input_term, search_formula=input_search_formula,
                         case_sensitive=input_case_sensitive, match_entire_val=input_match_entire_val,
                         sheet_name_contains=input_sheet_name_contains, sheet_name_excludes=input_sheet_name_excludes,
                         workbook_name_contains=input_workbook_name_contains,
                         workbook_name_excludes=input_workbook_name_excludes, recursive=input_recursive, colorize=True)

elif os.path.isfile(input_path):
    if not input_path.endswith(".xlsx"):
        raise Exception("Paths must be a directory or xlsx!")
    else:
        print("")
        results = search_workbook(path=input_path, term=input_term, search_formula=input_search_formula,
                                  case_sensitive=input_case_sensitive, match_entire_val=input_match_entire_val,
                                  sheet_name_contains=input_sheet_name_contains,
                                  sheet_name_excludes=input_sheet_name_excludes, colorize=True)
else:
    raise Exception("Paths must be a directory or xlsx! | {}".format(input_path))

print("")

print("Results!")
if len(results) == 0:
    print("There were none!")
else:
    for result in results:
        print(result)
