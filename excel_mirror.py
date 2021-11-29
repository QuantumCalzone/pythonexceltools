from excel_convert_workbook_to_csv import *
from excel_utils import *
import gc
import hashlib
import os
import shutil
import time

from pythonutils.os_utils import *

_verbose = False
_block_size = 65536  # The size of each read from the file


def move_csvs_and_hashes_to_match_workbooks(workbooks_dir_path, mirror_dir_path):
    if _verbose:
        print("move_csvs_and_hashes_to_match_workbooks ( workbooks_dir_path: {} , mirror_dir_path: {} )".format(
            workbooks_dir_path, mirror_dir_path))

    workbook_paths = get_all_in_dir(target_dir=workbooks_dir_path, full_path=True, recursive=True, include_dirs=False,
                                    include_files=True, must_end_in=".xlsx", include_hidden=False)
    workbook_paths = remove_temporary_workbooks(workbook_paths)
    workbook_rel_paths = []
    for workbook_path in workbook_paths:
        workbook_rel_path = workbook_path.replace(workbooks_dir_path, "")
        workbook_rel_path = workbook_rel_path.replace(".xlsx", "")
        workbook_rel_paths.append(workbook_rel_path)

    csv_dir_path = os.path.join(mirror_dir_path, "CSVs")
    csv_paths = get_all_in_dir(target_dir=csv_dir_path, full_path=True, recursive=True, include_dirs=False,
                               include_files=True, must_end_in=".csv", include_hidden=False)
    for csv_path in csv_paths:
        csv_path = get_parent_dir(csv_path)
        csv_rel_path = csv_path.replace(csv_dir_path, "")
        if csv_rel_path not in workbook_rel_paths:
            if os.path.exists(csv_path):
                shutil.rmtree(csv_path)
                if _verbose:
                    print("There is no workbook at csv_rel_path: {} so I'm deleting it at {}".format(csv_rel_path, csv_path))

    hashes_dir_path = os.path.join(mirror_dir_path, "Hashes")
    hashes_paths = get_all_in_dir(target_dir=hashes_dir_path, full_path=True, recursive=True, include_dirs=False,
                                  include_files=True, must_end_in=".hash", include_hidden=False)
    for hash_path in hashes_paths:
        hash_rel_path = hash_path.replace(hashes_dir_path, "")
        hash_rel_path = hash_rel_path.replace(".hash", "")
        if hash_rel_path not in workbook_rel_paths:
            os.remove(hash_path)
            if _verbose:
                print("There is no workbook at hash_rel_path: {} so I'm deleting it at {}".format(hash_rel_path, hash_path))


def get_hash_path(workbook_path, dir_path, hash_dir_path):
    # if _verbose:
    #     print("get_hash_path ( workbook_path: {} , dir_path: {} , hash_dir_path: {} )".format(workbook_path, dir_path,
    #                                                                                           hash_dir_path))

    hash_file_path = workbook_path.replace(dir_path, "")
    hash_file_path = hash_file_path.replace(".xlsx", ".hash")
    hash_file_path = hash_dir_path + hash_file_path

    return hash_file_path


def get_csv_path(workbook_path, dir_path, csv_dir_path):
    # if _verbose:
    #     print("get_csv_path ( workbook_path: {} , dir_path: {} , csv_dir_path: {} )".format(workbook_path, dir_path,
    #                                                                                           csv_dir_path))

    csv_file_path = workbook_path.replace(dir_path, "")
    csv_file_path = csv_file_path.replace(".xlsx", "")
    csv_file_path = csv_dir_path + csv_file_path

    return csv_file_path


def get_file_hash(file_path):
    # if _verbose:
    #     print("get_file_hash ( file_path: {} )".format(file_path))

    file_hash = hashlib.sha256()  # Create the hash object, can use something other than `.sha256()` if you wish

    with open(file_path, "rb") as file:  # Open the file to read it's bytes

        file_reader = file.read(_block_size)  # Read from the file. Take in the amount declared above

        while len(file_reader) > 0:  # While there is still data being read from the file

            file_hash.update(file_reader)  # Update the hash
            file_reader = file.read(_block_size)  # Read the next block from the file

    hexdigest = file_hash.hexdigest()

    return hexdigest


def save_hash_file(file_path, hash_file_path):
    if _verbose:
        print("save_hash_file ( workbook_path: {} , hash_file_path: {} )".format(file_path, hash_file_path))

    hex_digest = get_file_hash(file_path)

    ensure_dir_path_exists(os.path.dirname(hash_file_path))

    with open(hash_file_path, "w") as hash_file:
        hash_file.write(hex_digest)

    return hex_digest


def get_changed_workbooks(dir_path, hash_dir_path):
    if _verbose:
        print("get_changed_workbooks ( dir_path: {} , hash_dir_path: {} )".format(dir_path, hash_dir_path))

    changed_workbooks = []
    workbook_paths = get_all_in_dir(target_dir=dir_path, full_path=True, recursive=True, include_dirs=False,
                                    include_files=True, must_end_in=".xlsx", include_hidden=False)
    workbook_paths = remove_temporary_workbooks(workbook_paths)
    for workbook_path in workbook_paths:

        is_changed = False
        hash_file_path = get_hash_path(workbook_path, dir_path, hash_dir_path)

        if os.path.exists(hash_file_path):
            with open(hash_file_path, "r") as hash_file:
                hash_file_val = hash_file.read()
                workbook_hex_digest = get_file_hash(workbook_path)
                if hash_file_val != workbook_hex_digest:
                    is_changed = True
        else:
            is_changed = True

        if is_changed:
            changed_workbooks.append(workbook_path)

    return changed_workbooks


def mirror_workbooks(workbooks_dir_path, mirror_dir_path):
    if _verbose:
        print("mirror_workbooks ( workbooks_dir_path: {} , mirror_dir_path: {} )".format(workbooks_dir_path, mirror_dir_path))

    move_csvs_and_hashes_to_match_workbooks(workbooks_dir_path, mirror_dir_path)

    hashes_dir_path = os.path.join(mirror_dir_path, "Hashes")
    csv_dir_path = os.path.join(mirror_dir_path, "CSVs")

    changed_workbook_paths = get_changed_workbooks(workbooks_dir_path, hashes_dir_path)
    changed_workbook_paths_len = len(changed_workbook_paths)
    i = 0
    memory_object_count = len(gc.get_objects())
    for changed_workbook_path in changed_workbook_paths:

        last_memory_object_count = memory_object_count
        memory_object_count = len(gc.get_objects())
        print('last_memory_object_count: {} , object_count: {}, diff: {}'.format(
            ('{:,}'.format(last_memory_object_count)),
            ('{:,}'.format(memory_object_count)),
            (memory_object_count - last_memory_object_count))
        )

        i += 1
        print("Mirroring... {}%".format((float(i)/float(changed_workbook_paths_len))))
        hash_file_path = get_hash_path(changed_workbook_path, workbooks_dir_path, hashes_dir_path)
        save_hash_file(changed_workbook_path, hash_file_path)

        csv_export_path = get_csv_path(changed_workbook_path, workbooks_dir_path, csv_dir_path)
        convert_workbook(workbook_path=changed_workbook_path, export_path=csv_export_path)

        time.sleep(0.1)


def hash_all_workbooks(dir_path, hash_dir_parent_path):
    if _verbose:
        print("hash_all_workbooks ( dir_path: {} , hash_dir_parent_path: {} )".format(dir_path, hash_dir_parent_path))

    hash_dir_path = os.path.join(hash_dir_parent_path, "Hashes")

    workbook_paths = get_all_in_dir(target_dir=dir_path, full_path=True, recursive=True, include_dirs=False,
                                    include_files=True, must_end_in=".xlsx", include_hidden=False)
    workbook_paths = remove_temporary_workbooks(workbook_paths)
    workbook_paths_len = len(workbook_paths)
    i = 0
    for workbook_path in workbook_paths:
        i += 1
        print("Hashing... {}%".format((float(i)/float(workbook_paths_len))))
        hash_file_path = get_hash_path(workbook_path, dir_path, hash_dir_path)
        save_hash_file(workbook_path, hash_file_path)
