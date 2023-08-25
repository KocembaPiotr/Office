import os
import re
import time
import shutil
import glob


def files_copy_directory(path_source, path_destination, filter_re=None) -> int:
    """
    Function to copy all files from one directory to another.
    :param path_source: path to folder from files will be copied
    :param path_destination: path to folder where files will be copied
    :param filter_re: optional filter to indicate what files should be copied
    :return: 1 - as success, 0 - as failure
    """
    try:
        files = os.listdir(path_source)
        for fname in files:
            if re.search(filter_re, fname) or filter is None:
                shutil.copy2(os.path.join(path_source, fname), path_destination)
        return 1
    except OSError:
        return 0


def files_delete_directory(directory_path: str) -> int:
    """
    Function to delete all files in indicated directory.
    :param directory_path: path to indicated folder
    :return: 1 - as success, 0 - as failure
    """
    try:
        with os.scandir(directory_path) as entries:
            for entry in entries:
                if entry.is_file():
                    os.unlink(entry.path)
        return 1
    except OSError:
        return 0


def files_download_wait(directory: str, file_name: str, timeout_seconds: int,
                        nfiles: int = None) -> None:
    """
    Function to wait for files which are downloading.
    :param directory: path to folder where file will be downloaded
    :param file_name: name of the downloading file
    :param timeout_seconds: time computer will be wait if something go wrong
    :param nfiles: number of files
    :return: None
    """
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout_seconds:
        time.sleep(1)
        dl_wait = False
        files = os.listdir(directory)
        if nfiles and len(files) != nfiles:
            dl_wait = True
        if file_name not in files:
            dl_wait = True
        for fname in files:
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds


def files_list(folder: str, file_filter: str = None) -> list:
    if filter is None:
        return glob.glob(folder)
    else:
        return glob.glob(folder + file_filter)


def file_copy(path_source, path_destination) -> int:
    """
    Function to copy single file from one directory to another.
    :param path_source: path to folder from files will be copied
    :param path_destination: path to folder where files will be copied
    :return: 1 - as success, 0 - as failure
    """
    try:
        file_delete(path_destination)
        shutil.copy2(path_source, path_destination)
        return 1
    except OSError:
        return 0


def file_delete(path: str) -> None:
    """
    Function to delete proper file if it exists.
    :param path: path to proper file
    :return: None
    """
    if os.path.isfile(path):
        os.remove(path)


def folder_delete(dir_path: str) -> None:
    """
    Function to delete whole folder with all items inside.
    :param dir_path: path to folder
    :return: None
    """
    if os.path.exists(dir_path) and os.path.isdir(dir_path):
        shutil.rmtree(dir_path)


def folder_copy(dir_source_path: str, dir_destination_path: str) -> None:
    """
    FUnction to copy whole folder from one directory to another.
    :param dir_source_path: folder source path
    :param dir_destination_path: folder destination path
    :return: None
    """
    shutil.copytree(dir_source_path, dir_destination_path, dirs_exist_ok=True)