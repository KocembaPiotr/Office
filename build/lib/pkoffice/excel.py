import os
import sys
import time
import webbrowser
import xlwings as xw
from win32com.universal import com_error


def open_excel_sharepoint(file_path: str, file_name: str, time_limit: int = 20) -> xw.Book:
    """
    Function to open Excel on SharePoint and obtain its instance. file_path should have prefix: ms-excel:ofe|u|
    :param file_path: SharePoint path to Excel file with prefix: ms-excel:ofe|u|
    :param file_name: Excel file name
    :param time_limit: time limit to wait to open Excel on desktop app
    :return: xlwings book to further process
    """
    webbrowser.open(file_path)
    time.sleep(time_limit)
    return xw.Book(file_name)


def close_excel_instances(time_wait: int = 0) -> int:
    """
    Function will close all opened MS Excel instances.
    :param time_wait: wait before remove all Excel instances
    :return: None
    """
    time.sleep(time_wait)
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
        sh.api.ListObjects(table_name).Refresh()
    except Exception as e:
        print(e)


def file_columns_autofit(report_path: str, sheet: str) -> None:
    """
    Function to fit columns in Excel file automatically
    :param report_path: path to Excel file
    :param sheet: Sheet name of Excel file
    :return: None
    """
    wb = xw.Book(report_path)
    sh = wb.sheets(sheet)
    sh.autofit("columns")
    wb.save()
    wb.app.quit()
    close_excel_instances(3)
