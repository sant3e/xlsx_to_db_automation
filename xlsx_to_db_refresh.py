# in terminal: python -m pip install pywin32

import os
import numpy as np
import pandas as pd
import psycopg2
import shutil
import openpyxl
import win32com.client as win32
import time
from psycopg2.pool import ThreadedConnectionPool

start_time = time.time()


def bring_files(src_dir, dest_dir):  # Make a copy of working files into a temp dir (preserve connection)
    # set working files
    src_files = os.listdir(src_dir)

    # bring the copies in the working directory
    for file in src_files:
        full_file_name = os.path.join(src_dir, file)
        if os.path.isfile(full_file_name):
            shutil.copy(full_file_name, dest_dir)

    return


def xls_files(dest_dir):  # Find and isolate xlsx files
    xls_files = []
    for file in os.listdir(dest_dir):
        if file.endswith('.xlsx'):
            xls_files.append(file)

    return xls_files


def open_close_as_excel(files_path):   # Open xlsx', refreshes data from connection, save&close
    xlapp = win32.DispatchEx('Excel.Application')
    xlapp.DisplayAlerts = False
    xlapp.Visible = True

    xlbook = xlapp.Workbooks.Open(files_path)
    xlbook.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    xlbook.Save()
    xlbook.Close()
    xlapp.Quit()

    del xlbook
    del xlapp


def create_df(dest_dir, xls_files):    # Create Pandas DFs from the XLSX files
    data_path = dest_dir + '/'
    # loop through the files and create the dataframes
    df = {}
    for file in xls_files:
        try:
            df[file] = pd.read_excel(data_path+file)
        except UnicodeDecodeError:
            df[file] = pd.read_excel(data_path+file, encoding="ISO-8859-1")

    return df


def clean_tbl_name(filename):    # Preparing files &tables (compliant to SQL standards)
    # clean XLS file names
    clean_tbl_name = filename.lower().replace(" ", "_").replace("?", "").replace("$", "").replace("-", "_") \
        .replace(r"/", "_").replace("\\", "_").replace("%", "").replace(")", "").replace(r"(", "")

    # create table names (we need to remove '.xlsx' from the table names, so that SQL can name properly)
    tbl_name = '{0}'.format(clean_tbl_name.split('.')[0])

    return tbl_name


def clean_colname(dataframe):    # Preparing header row, data types and schema (compliant to SQL standards)
    # clean headers
    dataframe.columns = [x.lower().replace(" ", "_").replace("?", "").replace("$", "").replace("-", "_")
                             .replace(r"/", "_").replace("\\", "_").replace("%", "").replace(")", "")
                             .replace(r"(", "") for x in dataframe.columns]

    # replacement dictionary that maps pandas dtypes to proper sql dtypes
    replacements = {
        'object': 'text',
        'float64': 'float',
        'int64': 'int',
        'datetime64[ns]': 'timestamp',
        'timedelta64[ns]': 'text'
    }

    # table schema
    col_str = ", ".join(
        "{} {}".format(n, d) for (n, d) in zip(dataframe.columns, dataframe.dtypes.replace(replacements)))

    return col_str, dataframe.columns


def upload_to_db(host, dbname, user, password, tbl_name, col_str, file, dataframe, dataframe_columns):
    # Opening the database connection
    connection = psycopg2.connect("host=%s dbname=%s user=%s password=%s" % (host, dbname, user, password))
    cursor = connection.cursor()
    print('Opened database successfully')

    # Dropping tables with same name
    cursor.execute("DROP TABLE IF EXISTS %s;" % tbl_name)

    # Creating tables by passing the table names and table schemas
    cursor.execute("CREATE TABLE %s (%s);" % (tbl_name, col_str))
    print('{0} was created successfully'.format(tbl_name))

    # inserting values to table

    # Saving and loading pandas DFs into a CSV
    dataframe.to_csv(file, header=dataframe_columns, index=False, encoding='utf8')
    my_file = open(file, encoding='utf8')  # we save it as an object
    print('file opened in memory')

    # Uploading to DB (using copy_expert method for faster run-times)
    sql_statement = """
    COPY %s FROM STDIN WITH
    CSV
    HEADER
    """

    cursor.copy_expert(sql=sql_statement % tbl_name, file=my_file)
    print('file copied to db')

    # Granting access to all users; committing and closing
    cursor.execute("grant SELECT on TABLE %s to public" % tbl_name)
    connection.commit()
    cursor.close()
    print('table {0} imported to db completed'.format(tbl_name))

    return


# main program

# DB credentials
host = '<host_location>'
dbname = '<database_name>'
user = "<user_name>"
password = "<password>"


# set working directories
src_dir = os.getcwd() + '/' + 'original'
dest_dir = os.getcwd() + '/' + 'source'

## copies of working xlsx files
# after running upload_to_db xlsx get screwed
# so, we need to overwrite them each time when running script
bring_files(src_dir, dest_dir)

# refresh the xlsx files (bound to sharepoint lists; connections are created by user using Power Query)
for file in os.listdir(dest_dir):
    xlsx_path = os.path.join(dest_dir, file)
    open_close_as_excel(xlsx_path)

# create the dataframes for tables
xls_df = xls_files(dest_dir)
df = create_df(dest_dir, xls_df)

# iterating through xlsx files
for k in xls_df:
    # call dataframe
    dataframe = df[k]

    # clean table name
    tbl_name = clean_tbl_name(k)

    # clean column names
    col_str, dataframe.columns = clean_colname(dataframe)

    # upload to db
    upload_to_db(host,
                 dbname,
                 user,
                 password,
                 tbl_name,
                 col_str,
                 file=k,
                 dataframe=dataframe,
                 dataframe_columns=dataframe.columns)

print("Script runs for: ", time.time()-start_time)
