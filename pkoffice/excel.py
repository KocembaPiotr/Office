import os
import sys
from win32com.universal import com_error


def close_excel_instances() -> int:
    """
    Function will close all opened MS Excel instances.
    :return: None
    """
    os.system(f'taskkill /F /IM Excel.exe')


def filters_clean_filter(wb, sh) -> None:
    """
    Function to clean Excel filters but keep them as they were.
    :param wb: workbook variable
    :param sh: sheet variable
    :return: None
    """
    if sh.api.AutoFilterMode:
        sh.api.AutoFilter.ShowAllData()
    try:
        wb.api.Names.Item("_FilterDatabase").Delete()
    except com_error:
        print(sys.exc_info())


def filters_remove_filter(wb, sh) -> None:
    """
    Function to remove all Excel filters on indicated sheet.
    :param wb: workbook variable
    :param sh: sheet variable
    :return: None
    """
    if sh.api.AutoFilterMode:
        sh.api.AutoFilterMode = False
    try:
        wb.api.Names.Item("_FilterDatabase").Delete()
    except com_error:
        print(sys.exc_info())
