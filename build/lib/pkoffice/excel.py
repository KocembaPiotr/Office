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


def refresh_table_pivot(sh, table_pivot_name: str) -> None:
    """
    Function to refresh pivot table in Excel.
    :param table_pivot_name: name of pivot table which will be refreshed
    :param sh: sheet variable
    :return: None
    """
    try:
        sh.api.PivotTables(table_pivot_name).RefreshTable()
    except Exception as e:
        print(e)


def refresh_table(sh, table_name: str) -> None:
    """
    Function to refresh table in Excel.
    :param table_name: name of pivot table which will be refreshed
    :param sh: sheet variable
    :return: None
    """
    try:
        sh.api.ListObjects(table_name).Refresh
    except Exception as e:
        print(e)
