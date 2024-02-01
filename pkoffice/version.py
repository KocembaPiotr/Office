import re
import file


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
    local_functions_version = version_read(file_local_path, version_pattern)
    server_functions_version = version_read(file_server_path, version_pattern)
    if local_functions_version != server_functions_version:
        file.file_copy(file_server_path, file_local_path)