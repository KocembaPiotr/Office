import os
import sys
import time
import webbrowser
import pandas as pd
import xlwings as xw
import win32com.client
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


def close_columns_autofit(report_path: str, sheet: str) -> None:
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


def filters_clean_filter(wb: xw.Book, sh: xw.sheets) -> None:
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


def filters_remove_filter(wb: xw.Book, sh: xw.sheets) -> None:
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


def refresh_table_pivot(sh: xw.sheets, table_pivot_name: str) -> None:
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


def refresh_table(sh: xw.sheets, table_name: str) -> None:
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


def refresh_all_tables_without_open(file: str) -> None:
    """
    Function to refresh table in Excel without it open.
    :param file: path to Excel file which need to be refreshed
    :return: None
    """
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(file)
    wb.RefreshAll()
    wb.Save()
    wb.Close(SaveChanges=False)
    excel.Quit()


def refresh_macro_without_open(file: str, macro: str) -> None:
    """
    Function to refresh macro in Excel without it open.
    :param file: path to Excel file which need to be refreshed
    :param macro: macro name which is located in Excel file
    :return: None
    """
    excel = win32com.client.Dispatch("Excel.Application")
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(file)
        excel.Application.Run(macro)
        workbook.Close(SaveChanges=True)
    except Exception as e:
        print(e)
    finally:
        excel.Quit()


def create_table(df: pd.DataFrame, sh: xw.sheets, table_start_range: str, table_name: str) -> None:
    """
    Function to create Excel table base on python data
    :param df: pandas dataframe
    :param sh: Excel sheet
    :param table_start_range: start of the range
    :param table_name: table name
    :return: None
    """
    for table in sh.api.ListObjects:
        if table.Name == table_name:
            table.Delete()
    sh[table_start_range].options(pd.DataFrame, header=1, index=False, expand='table').value = df
    table_range = sh.range(table_start_range).expand('table')
    sh.api.ListObjects.Add(1, sh.api.Range(table_range.address), None, 1).Name = table_name


def create_table_win32(df: pd.DataFrame, excel_path: str, excel_sheet: str, table_start_range: str, table_name: str) -> None:
    """
    Function to create Excel table base on python data and win32com
    :param df: pandas dataframe
    :param excel_path: Excel workbook
    :param excel_sheet: Excel sheet
    :param table_start_range: start of the range
    :param table_name: table name
    :return: None
    """
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks(excel_path)
    sheet = workbook.Sheets(excel_sheet)
    for table in sheet.ListObjects:
        if table.Name == table_name:
            table.Delete()
    sheet_range = f"{table_start_range}:{chr(64 + df.shape[1]) + str(df.shape[0]+1)}"
    sheet.Range(sheet_range).Value = [df.columns.values.tolist(), *df.values.tolist()]
    sheet.ListObjects.Add(1, sheet.Range(sheet_range), None, 1).Name = table_name


def df_to_excel(df: pd.DataFrame, file_path: str, sheet_name: str = 'Sheet1',
                cond_format: dict = None, header_format: dict = None) -> None:
    """
    Function to save dataframe to excel with autofit columns
    :param df: pandas dataframe
    :param file_path: path where file will be saved
    :param sheet_name: name of Excel sheet
    :param cond_format: optional conditional formatting
    :param header_format: optional header formatting
    :return: None
    example of dictionaries:
    cond_format={'range': 'F2:F10000', 'type': 'cell', 'criteria': '>', 'value': 4000,
                 'format': {'bg_color': '#D9D9D9'}},
    header_format={"bg_color": "#00D100", "border": 2, "bold": True}
    """
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        for column in df:
            column_length = max(df[column].astype(str).map(len).max(), len(column)) * 1.2
            col_idx = df.columns.get_loc(column)
            writer.sheets[sheet_name].set_column(col_idx, col_idx, column_length)
        if cond_format is not None:
            cond_format_range = cond_format.pop('range')
            cond_format['format'] = writer.book.add_format(cond_format['format'])
            writer.sheets[sheet_name].conditional_format(cond_format_range, cond_format)
        if header_format is not None:
            for idx2, col in enumerate(df.columns):
                writer.sheets[sheet_name].write(0, idx2, col, writer.book.add_format(header_format))


def df_to_excel_list(df_list: list[pd.DataFrame], file_path: str, sheet_list: list) -> None:
    """
    Function to save dataframe to excel with autofit columns
    :param df_list: list of pandas dataframes
    :param file_path: path where file will be saved
    :param sheet_list: list of names of Excel sheets
    :return: None
    """
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        for df, sheet_name in zip(df_list, sheet_list):
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            for column in df:
                column_length = max(df[column].astype(str).map(len).max(), len(column)) * 1.2
                col_idx = df.columns.get_loc(column)
                writer.sheets[sheet_name].set_column(col_idx, col_idx, column_length)


def copy_column_win32(excel_path: str, excel_sheet: str, excel_column: str) -> pd.DataFrame:
    """
    Function to copy data from indicated column to pandas dataframe
    :param excel_path: Excel workbook
    :param excel_sheet: Excel sheet
    :param excel_column: Indicated Excel column to copy data
    :return: pandas dataframe
    """
    df = pd.DataFrame()
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks(excel_path)
    try:
        sheet = workbook.Sheets(excel_sheet)
        column_data = sheet.Range(f"{excel_column}:{excel_column}0000").Value
        df = pd.DataFrame(column_data[1:], columns=[column_data[0][0]])
    finally:
        return df