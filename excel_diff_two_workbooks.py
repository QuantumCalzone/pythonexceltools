from excel_convert_workbook_to_csv import *
from excel_diff_two_csvs import *
from pythonutils.os_utils import *

_verbose = False


def workbook_diff(workbook_1_path, workbook_2_path, diff_output_dir, only_show_diff):
    if _verbose:
        print("workbook_diff ( workbook_1_path: {} , workbook_2_path: {} , diff_output_dir: {} , only_show_diff: {} )".format(workbook_1_path, workbook_2_path, diff_output_dir, only_show_diff))

    diff_output_dir = make_unique_time_folder_at_path(diff_output_dir)

    workbook_1_name = get_file_name_from_path(workbook_1_path)
    workbook_2_name = get_file_name_from_path(workbook_2_path)

    if workbook_1_name == workbook_2_name:
        workbook_1_name = "{}_A".format(workbook_1_name)
        workbook_2_name = "{}_B".format(workbook_2_name)

    csv_sources_path = os.path.join(diff_output_dir, "CSV Sources")
    csv_source_output_dir_workbook_1 = os.path.join(csv_sources_path, workbook_1_name)
    csv_source_output_dir_workbook_2 = os.path.join(csv_sources_path, workbook_2_name)

    convert_workbook(workbook_path=workbook_1_path, export_path=csv_source_output_dir_workbook_1)
    convert_workbook(workbook_path=workbook_2_path, export_path=csv_source_output_dir_workbook_2)

    workbook_1_csvs = get_all_in_dir(csv_source_output_dir_workbook_1, full_path=False, recursive=True, include_dirs=False,
                                     include_files=True, must_end_in=".csv", include_hidden=False)
    workbook_2_csvs = get_all_in_dir(csv_source_output_dir_workbook_2, full_path=False, recursive=True, include_dirs=False,
                                     include_files=True, must_end_in=".csv", include_hidden=False)

    diff_sheets_output_dir = os.path.join(diff_output_dir, "Diff")
    ensure_dir_path_exists(diff_sheets_output_dir)

    deleted_sheets = ""
    new_sheets = ""

    for workbook_2_sheet_name in workbook_2_csvs:
        if workbook_2_sheet_name in workbook_1_csvs:
            workbook_1_sheet_path = os.path.join(csv_source_output_dir_workbook_1, workbook_2_sheet_name)
            workbook_2_sheet_path = os.path.join(csv_source_output_dir_workbook_2, workbook_2_sheet_name)
            # print("workbook_1_sheet_path: {}".format(workbook_1_sheet_path))
            # print("workbook_2_sheet_path: {}".format(workbook_2_sheet_path))
            csv_diff(workbook_1_sheet_path, workbook_2_sheet_path, diff_sheets_output_dir, only_show_diff)
        else:
            new_sheets += "{}\n".format(workbook_2_sheet_name)

    for workbook_1_sheet_name in workbook_1_csvs:
        if workbook_1_sheet_name not in workbook_2_csvs:
            deleted_sheets += "{}\n".format(workbook_1_sheet_name)

    if deleted_sheets != "":
        deleted_sheets_file_name = "DeletedSheets {} -> {}.txt".format(workbook_1_name, workbook_2_name)
        deleted_sheets_file_path = os.path.join(diff_output_dir, deleted_sheets_file_name)
        print("deleted_sheets_file_path: {}".format(deleted_sheets_file_path))
        with open(deleted_sheets_file_path, "w") as deleted_sheets_file:
            deleted_sheets_file.write(deleted_sheets)

    if new_sheets != "":
        new_sheets_file_name = "NewSheets {} -> {}.txt".format(workbook_1_name, workbook_2_name)
        new_sheets_file_path = os.path.join(diff_output_dir, new_sheets_file_name)
        with open(new_sheets_file_path, "w") as new_sheets_file:
            new_sheets_file.write(new_sheets)
