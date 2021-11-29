import csv
import os
from pythonutils.os_utils import *

_verbose = False


def csv_reader_to_matrix(csv_reader):
    matrix = []

    for row in csv_reader:
        matrix.append(row)

    return matrix


def get_row_diff(row_1, row_2, only_show_diff, is_first_row):
    if _verbose:
        print("get_row_diff ( row_1: {} , row_2: {} , only_show_diff: {} , is_first_row: {} )".format(row_1, row_2, only_show_diff, is_first_row))

    get_row_diff_result = ""
    row_1_len = len(row_1)
    row_2_len = len(row_2)
    most_cols = row_1_len if row_1_len > row_2_len else row_2_len
    col_index = 0

    while col_index < most_cols:

        row_1_val = "OOB" if col_index >= row_1_len else row_1[col_index]
        row_2_val = "OOB" if col_index >= row_2_len else row_2[col_index]
        if not is_first_row and only_show_diff:
            get_row_diff_result += ""
            if row_1_val != row_2_val:
                get_row_diff_result += "{} -> {}".format(row_1_val, row_2_val)
        else:
            get_row_diff_result += "{} -> {}".format(row_1_val, row_2_val) if row_1_val != row_2_val else row_1_val

        col_index += 1

        if col_index < most_cols:
            get_row_diff_result += ','

    # print("get_row_diff_result: {}".format(get_row_diff_result))

    return get_row_diff_result


def csv_diff(csv_1_path, csv_2_path, diff_output_dir, only_show_diff):
    if _verbose:
        print("csv_diff ( csv_1_path: {} , csv_2_path: {} , diff_output_dir: {} , only_show_diff: {} )".format(csv_1_path, csv_2_path, diff_output_dir, only_show_diff))

    diff_output_csv_path = os.path.join(diff_output_dir, "{}-to-{}".format(get_file_name_from_path(csv_1_path), get_file_name_from_path(csv_2_path))) + ".csv"

    diff = ""

    with open(csv_1_path, "r") as csv_1_file, open(csv_2_path, "r") as csv_2_file:
        csv_1_reader = csv.reader(csv_1_file, delimiter=",", quotechar="|", skipinitialspace=True)
        csv_2_reader = csv.reader(csv_2_file, delimiter=",", quotechar="|", skipinitialspace=True)

        csv_1_matrix = csv_reader_to_matrix(csv_1_reader)
        csv_2_matrix = csv_reader_to_matrix(csv_2_reader)

        csv_1_matrix_rows_len = len(csv_1_matrix)
        csv_2_matrix_rows_len = len(csv_2_matrix)

        row_index = 0
        most_rows = csv_1_matrix_rows_len if csv_1_matrix_rows_len > csv_2_matrix_rows_len else csv_2_matrix_rows_len

        while row_index < most_rows:
            csv_1_row = [] if row_index >= csv_1_matrix_rows_len else csv_1_matrix[row_index]
            csv_2_row = [] if row_index >= csv_2_matrix_rows_len else csv_2_matrix[row_index]

            diff += get_row_diff(csv_1_row, csv_2_row, only_show_diff, row_index == 0)

            row_index += 1

            if row_index < most_rows:
                diff += '\n'

        with open(diff_output_csv_path, "w") as diff_output_csv_file:
            diff_output_csv_file.write(diff)
