import re
import sqlalchemy as sql
import pandas as pd
import numpy as np
from typing import Literal
from datetime import datetime
from pkoffice import file

TMP_FILE = 'tmp.csv'


class SqlDB:
    """
    Class to manage sql database connection.
    """
    def __init__(self, server: str, database: str, driver: str):
        self.engine = sql.create_engine(f"mssql+pyodbc://{server}/"
                                        f"{database}?driver={driver}",
                                        fast_executemany=True)
        self.database = database
        self.process_time_beg = datetime.now()
        self.process_time_end = datetime.now()
        self.df = None
        self.table = None
        self.flag_commit = None

    def download_data(self, query: str) -> pd.DataFrame:
        """
        Method to download data from database according to provided query.
        :param query: SQL query in string format
        :return: pandas.Dataframe
        """
        with self.engine.begin() as conn:
            return pd.read_sql(sql=sql.text(query), con=conn)

    def execute_query(self, query: str) -> None:
        """
        Method to execute query
        :param query: SQL query in string format
        :return: None
        """
        with self.engine.begin() as conn:
            conn.execute(sql.text(query))

    def upload_data(self, df: pd.DataFrame, table_name: str, chunksize: int,
                    if_exists: Literal["new", "replace", "append"] = 'replace',
                    log_table: str = None) -> None:
        """
        Method to upload pandas dataframe to database.
        :param df: pandas dataframe with data to upload
        :param table_name: name of table without []
        :param chunksize: size of single batch to upload
        :param if_exists: replace - drop/create, append - insert at the end,
               new - delete and insert
        :param log_table: name of log table to provide there upload parameters
        :return: None
        """
        try:
            if if_exists == 'new':
                with self.engine.begin() as conn:
                    conn.execute(sql.text(f'Delete from dbo.[{table_name}]'))
                    df.to_sql(table_name, con=conn, if_exists='append',
                              index=False, chunksize=chunksize, method='multi',
                              schema='dbo')
            else:
                df.to_sql(table_name, con=self.engine, if_exists=if_exists,
                          index=False, chunksize=chunksize, method='multi', schema='dbo')
            self.flag_commit = True
        except Exception as e:
            self.flag_commit = False
            print(e)
        self.process_time_end = datetime.now()
        self.df = df
        self.table = table_name
        if log_table is not None:
            self.upload_log(log_table, self.upload_parameters())
            self.df = None

    def upload_data_mass(self, df: pd.DataFrame, table_name: str, chunksize: int = 300,
                         flag_delete_data: bool = True, log_table: str = None) -> None:
        """
        Method to upload large pandas dataframe to database in mass in multiple chunks.
        :param df: pandas dataframe with data to upload
        :param table_name: name of table without []
        :param chunksize: size of single batch to upload
        :param flag_delete_data: flag to indicate if user would like first delete data from table
        :param log_table: name of log table to provide there upload parameters
        :return: None
        """
        def upload_data_mass_single(df_sliced):
            sql_query_batch = ''
            for val in df_sliced.values:
                sql_query = f"Insert into dbo.[{table_name}] Values ({','.join(val)})" \
                    .replace("'nan'", 'null').replace("nan", "null")
                sql_query_batch = f'{sql_query_batch} {sql_query}'
            conn.execute(sql.text(sql_query_batch))

        def replace_string(x):
            return re.sub(':', ': ', re.sub("'", "", str(x)))

        for i in df:
            if df[i].dtype == 'datetime64[ns]':
                df[i] = df[i].dt.strftime('%Y-%m-%d %H:%M:%S')
                df[i] = df[i].apply(lambda x: f"'{x}'")
                df[i] = df[i].astype('string')
            elif df[i].dtype == 'float64':
                df[i] = df[i].apply(lambda x: f"{x:.4f}")
                df[i] = df[i].astype('string')
            elif df[i].dtype == 'int64':
                df[i] = df[i].apply(lambda x: f"{x}")
                df[i] = df[i].astype('string')
            else:
                df[i] = df[i].apply(lambda x: f"'{replace_string(x)}'")
                df[i] = df[i].astype('string')
        try:
            with self.engine.begin() as conn:
                if flag_delete_data:
                    conn.execute(sql.text(f'Delete from dbo.[{table_name}]'))
                beg = 0
                for i in range(0, len(df), chunksize):
                    if i != 0:
                        upload_data_mass_single(df[beg:i])
                        beg = i
                if len(df) >= beg:
                    upload_data_mass_single(df[beg:])
            self.flag_commit = True
        except Exception as e:
            print(e)
            self.flag_commit = False
        self.process_time_end = datetime.now()
        self.df = df
        self.table = table_name
        if log_table is not None:
            self.upload_log(log_table, self.upload_parameters())
            self.df = None

    def upload_bulk(self, df: pd.DataFrame, table_name: str, server_folder: str,
                    flag_delete_data: bool = True, log_table: str = None) -> None:
        """
        Method to upload data to database using INSERT BULK.
        :param df: pandas dataframe with data to upload
        :param table_name: name of table without []
        :param server_folder: path to folder which is visible to server to insert
        :param flag_delete_data: flag to indicate if user would like first delete data from table
        :param log_table: name of log table to provide there upload parameters
        :return: None
        """
        try:
            file_tmp = server_folder + TMP_FILE
            file.file_delete(file_tmp)
            df = df.replace(',', '', regex=True).replace('&', '', regex=True).\
                replace('"', '', regex=True).replace("'", "", regex=True)
            df.to_csv(file_tmp, index=False, header=False)
            with self.engine.begin() as conn:
                if flag_delete_data:
                    conn.execute(sql.text(f'Delete from dbo.[{table_name}]'))
                conn.execute(sql.text(f"""
                                        Bulk Insert {self.database}.dbo.[{table_name}] From '{file_tmp}' 
                                        WITH (FIELDTERMINATOR = ',')
                                     """))
            file.file_delete(file_tmp)
            self.flag_commit = True
        except Exception as e:
            self.flag_commit = False
            print(e)
        self.process_time_end = datetime.now()
        self.df = df
        self.table = table_name
        if log_table is not None:
            self.upload_log(log_table, self.upload_parameters())
            self.df = None

    def upload_parameters(self) -> None:
        """
        Method to return main process parameters.
        :return: [process date, process time, process duration,
                  dataframe records]
        """
        if self.df is not None:
            df_db_max = self.download_data(f"Select count(*) from dbo.[{self.table}]").values[0][0]
            df_max = self.df.count().max()
            df_check = 'Commit'
            if not self.flag_commit:
                df_check = 'RollBack'
            elif df_max != df_db_max:
                df_check = 'Fault'
            return [self.process_time_end.strftime("%Y-%m-%d"),
                    self.process_time_end.strftime("%H:%M:%S"),
                    self.table,
                    (self.process_time_end - self.process_time_beg).seconds,
                    (self.process_time_end - self.process_time_beg).seconds,
                    df_max, df_check]
        else:
            return [self.process_time_end.strftime("%Y-%m-%d"),
                    self.process_time_end.strftime("%H:%M:%S"),
                    (self.process_time_end - self.process_time_beg).seconds]

    def upload_log(self, table_name: str, log_value: list) -> None:
        """
        Method to upload log to database.
        :param table_name: log table name in SQL database without []
        :param log_value: list of variables to upload to log table
        :return: None
        """
        self.process_time_end = datetime.now()
        sql_values = [f"'{x}'" if type(x) == str else str(x) for x in log_value]
        sql_values = ','.join(sql_values)
        with self.engine.begin() as conn:
            conn.execute(
                f"""Insert into dbo.[{table_name}]
                Values ({sql_values})""")


def columns_str_max_len(df: pd.DataFrame) -> None:
    """
    Function to return max length of string columns multiply by 1.5.
    :param df: dataframe which will be uploaded to database
    :return: None
    """
    col_string = [i for i in df if df[i].dtype in ('object', 'string')]
    df[col_string] = df[col_string].astype('string')
    df_col_max_len = [{i: int(np.ceil(1.5 * df[i].str.len().max()))} for i in col_string]
    for column in df_col_max_len:
        print(column)


def parse_date_from_number(df: pd.DataFrame) -> pd.DataFrame:
    """
    Function to parse date from ordinal numbers.
    :param df: dataframe with only date columns
    :return: dataframe with converted columns
    """
    for col in df.columns:
        df[col] = df[col].apply(
            lambda x: datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(x) - 2))
    return df