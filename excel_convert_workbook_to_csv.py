from excel_utils import *
import openpyxl
from pythonutils.colors_utils import *
from pythonutils.os_utils import *
from pythonutils.str_utils import *

_verbose = False
_super_verbose = False


def convert_dir(dir_path, export_path, recursive):
    if _verbose:
        print(f'convert_dir ( dir_path: {dir_path} , export_path: {export_path} , recursive: {recursive} )')

    # matching workbook filter will not support case sensitivity
    workbook_paths = get_all_in_dir(
        target_dir=dir_path, full_path=True, recursive=recursive,
        include_dirs=False, include_files=True, must_end_in='.xlsx')
    workbook_paths = remove_temporary_workbooks(workbook_paths)

    workbook_path_count = len(workbook_paths)
    workbook_paths_searched = 0
    for workbook_path in workbook_paths:
        workbook_paths_searched += 1
        workbook_name = get_file_name_from_path(workbook_path)

        if workbook_name.startswith('~'):
            print(f'{num_to_comma_str(workbook_paths_searched)}/{num_to_comma_str(workbook_path_count)} '
                  f'skipping workbook: {get_yellow(workbook_name)}')
        else:
            print('{}/{} converting workbook: {}'.format(num_to_comma_str(
                workbook_paths_searched), num_to_comma_str(workbook_path_count), get_yellow(workbook_name)))

            target_export_path = export_path

            if recursive:
                target_export_path_sub = workbook_path.replace(dir_path, '')
                target_export_path_sub = target_export_path_sub.replace('{}.xlsx'.format(workbook_name), '')
                if target_export_path_sub != '':
                    # remove first char from string
                    target_export_path_sub = target_export_path_sub[1:]
                    target_export_path = os.path.join(export_path, target_export_path_sub)

            convert_workbook(workbook_path=workbook_path, export_path=target_export_path)


def convert_workbook(workbook_path, export_path):
    log = 'convert_workbook ( workbook_path: {} , export_path: {} )'.format(workbook_path, export_path)
    if _verbose:
        print(log)

    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    workbook_name = get_file_name_from_path(workbook_path, with_extension=False)
    ensure_dir_path_exists(export_path)

    sheet_names = workbook.get_sheet_names()

    for sheet_name in sheet_names:
        if _verbose:
            print('{} | sheet_name: {}'.format(log, workbook_path, export_path, sheet_name))

        sheet = workbook.get_sheet_by_name(sheet_name)

        rows = sheet.rows

        csv_path = os.path.join(export_path, '{}.csv'.format(sheet_name))
        csv = open(csv_path, 'w+')

        for row in rows:
            row_list = list(row)
            row_list_length = len(row_list)
            row_list_range = range(row_list_length)

            for i in row_list_range:
                val = row_list[i].value
                if type(val) is str:
                    val = '' if val is None else val.encode('ascii', 'ignore').decode('ascii')
                else:
                    val = '' if val is None else str(val)
                if i == row_list_length - 1:
                    csv.write(val)
                else:
                    csv.write(val + ',')
            csv.write('\n')
        csv.close()

    workbook.close()
