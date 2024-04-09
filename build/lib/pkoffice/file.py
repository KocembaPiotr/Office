import datetime
import os
import re
import time
import shutil
import glob


def files_copy_directory(path_source, path_destination, filter_re=None) -> int:
    """
    Function to copy all files from one directory to another
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
    Function to delete all files in indicated directory
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
    Function to wait for files which are downloading
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


def files_list(folder: str, file_filter: str = '*') -> list:
    """
    Method to list files from indicated folder using filter
    :param folder: indicated folder to search
    :param file_filter: filter which files to choose
    :return: None
    """
    return glob.glob(folder + file_filter)


def file_copy(path_source, path_destination, folder_creation=False) -> int:
    """
    Function to copy single file from one directory to another
    :param folder_creation: flag to check if folder exists and if not create folder for destination purpose
    :param path_source: path to folder from files will be copied
    :param path_destination: path to folder where files will be copied
    :return: 1 - as success, 0 - as failure
    """
    try:
        if folder_creation:
            dir_name, _ = os.path.split(path_destination)
            os.makedirs(dir_name, exist_ok=True)
        file_delete(path_destination)
        shutil.copy2(path_source, path_destination)
        return 1
    except OSError as e:
        print(e)
        return 0


def file_delete(path: str) -> None:
    """
    Function to delete proper file if it exists
    :param path: path to proper file
    :return: None
    """
    if os.path.isfile(path):
        os.remove(path)


def file_exists(path: str) -> bool:
    """
    Function to check if file exists in indicated location
    :param path: path to proper file
    :return: True - if exists, False if it does not exist
    """
    if os.path.isfile(path):
        return True
    else:
        return False


def file_date_creation(file_path: str) -> datetime.datetime:
    return datetime.datetime.fromtimestamp(os.path.getctime(file_path))


def file_date_modification(file_path: str) -> datetime.datetime:
    return datetime.datetime.fromtimestamp(os.path.getmtime(file_path))


def file_base_name(path: str, separator: str = '\\') -> str:
    """
    Function to return base file name from path
    :param path: full path with file name
    :param separator: separator used in split
    :return: file name
    """
    head, *_, tail = path.split(separator)
    return tail


def file_name_change(old_name: str, new_name: str) -> None:
    """
    Function to change name of the file
    :param old_name: path to file with old name
    :param new_name: path to file with new name
    :return: None
    """
    os.rename(old_name, new_name)


def file_content_read(path: str) -> list:
    """
    Function to read data from file
    :param path: path to file
    :return: content of the file
    """
    with open(path, 'r') as f:
        return f.readlines()


def file_content_write(path: str, content_new: list) -> None:
    """
    Function to write data to file
    :param path: path to file
    :param content_new: new content to write
    :return: None
    """
    with open(path, 'w') as f:
        f.writelines(content_new)


def file_content_remove_lines(file_path: str, cond_in: str) -> None:
    """
    Function to remove proper lines in file
    :param file_path: path to file
    :param cond_in: parameter which is in line to be removed
    :return: None
    """
    try:
        with open(file_path, 'r') as file:
            lines = file.readlines()
        with open(file_path, 'w') as file:
            for line in lines:
                if cond_in not in line:
                    file.write(line)
    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")


def file_content_replace(file_path: str, old_char: str, new_char: str) -> None:
    """
    Function to replace proper char in file
    :param file_path: path to file
    :param old_char: old char to be replaced
    :param new_char: new char
    :return: None
    """
    try:
        with open(file_path, 'r') as file:
            lines = file.read()
        modified_lines = lines.replace(old_char, new_char)
        with open(file_path, 'w') as file:
            file.write(modified_lines)
    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")


def folder_delete(dir_path: str) -> None:
    """
    Function to delete whole folder with all items inside
    :param dir_path: path to folder
    :return: None
    """
    if os.path.exists(dir_path) and os.path.isdir(dir_path):
        shutil.rmtree(dir_path)


def folder_copy(dir_source_path: str, dir_destination_path: str) -> None:
    """
    FUnction to copy whole folder from one directory to another
    :param dir_source_path: folder source path
    :param dir_destination_path: folder destination path
    :return: None
    """
    shutil.copytree(dir_source_path, dir_destination_path, dirs_exist_ok=True)


def folder_create(folder_path: str) -> None:
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


def version_read(file_path: str, version_pattern: str) -> str:
    """
    :param file_path: path to file where version is put
    :param version_pattern: pattern to read version number
    :return: version statement
    """
    try:
        with open(file_path, 'r') as f:
            first_line = f.readline()
            return re.sub(version_pattern, '', first_line)
    except Exception as e:
        print(e)
        return ''


def version_update(file_local_path: str, file_server_path: str,
                   version_pattern: str) -> None:
    """
    Method to update version to new one.
    :param file_local_path: path to file on local computer
    :param file_server_path: path to file on server repository
    :param version_pattern: pattern to extract proper version number
    :return: None
    """
    if file_exists(file_local_path):
        local_functions_version = version_read(file_local_path, version_pattern)
        server_functions_version = version_read(file_server_path, version_pattern)
        if local_functions_version != server_functions_version:
            file_copy(file_server_path, file_local_path)
    else:
        file_copy(file_server_path, file_local_path)