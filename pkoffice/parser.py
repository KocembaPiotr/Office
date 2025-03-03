import pandas as pd
import numpy as np
from datetime import datetime


def parse_to_date_from_number(df: pd.DataFrame,  column_names: list) -> pd.DataFrame:
    """
    Function to parse date from ordinal numbers
    :param df: pandas dataframe with data to convert
    :param column_names: list of columns which need to be converted
    :return: pandas dataframe with corrected types
    """
    for column_name in column_names:
        df[column_name] = df[column_name].apply(
            lambda x: datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(x) - 2))
    return df


def parse_to_date_from_str(df: pd.DataFrame,  column_names: list,
                           format_from: str = '%d.%m.%Y',
                           format_to: str = '%Y-%m-%d') -> pd.DataFrame:
    """
    Function to parse date from string
    :param df: pandas dataframe with data to convert
    :param column_names: list of columns which need to be converted
    :param format_from: date format in string
    :param format_to: desired date format
    :return: pandas dataframe with corrected types
    """
    for column_name in column_names:
        df[column_name] = pd.to_datetime(df[column_name], errors='coerce',
                                         format=format_from).dt.strftime(format_to)
    return df


def parse_to_float_from_time(df: pd.DataFrame, column_names: list) -> pd.DataFrame:
    """
    Function to convert data to float type
    :param df: pandas dataframe with data to convert
    :param column_names: column_names: list of columns which need to be converted
    :return: pandas dataframe with corrected types
    """
    def parse_to_float_from_time_func(time):
        total_seconds = time.hour * 3600 + time.minute * 60 + time.second
        return total_seconds / (24 * 60 * 60)
    for col in column_names:
        df[col] = pd.to_datetime(df[col], format='%H:%M:%S').dt.time
        df[col] = df[col].apply(parse_to_float_from_time_func)
    return df


def parse_to_float(df: pd.DataFrame, column_names: list) -> pd.DataFrame:
    """
    Function to convert data to float type
    :param df: pandas dataframe with data to convert
    :param column_names: list of columns which need to be converted
    :return: pandas dataframe with corrected types
    """
    df[column_names] = df[column_names].replace(' ', '', regex=True).replace(',', '.', regex=True).astype('float')
    return df


def parse_to_int_with_str_nan(df: pd.DataFrame, column_names: list) -> pd.DataFrame:
    """
    Function to convert data to int type with nulls
    :param df: pandas dataframe with data to convert
    :param column_names: list of columns which need to be converted
    :return: pandas dataframe with corrected types
    """
    df[column_names] = df[column_names].replace(' ', '', regex=True)
    for column_name in column_names:
        df[column_name] = df[column_name].fillna(-9999)
        df[column_name] = df[column_name].astype(int)
        df[column_name] = df[column_name].astype(str)
        df[column_name] = df[column_name].replace('-9999', np.nan)
    return df