import pandas as pd
import os
import datetime as dt
from log import logStatus
from tempfile import NamedTemporaryFile
import win32com.client as client
xlApp = client.Dispatch("Excel.Application")
pd.set_option('display.max_columns', None)


def get_current_month():
    """
    This function returns current month
    """
    return int(dt.date.today().strftime('%m'))


def get_current_day():
    """
    This function returns current day
    """
    return int(dt.date.today().strftime('%d'))


def get_current_year():
    """
    This function returns current year
    This function is used in FW_Validation.py
    """
    return int(dt.date.today().strftime('%Y'))


def month_diff(d1, d2):
    return (d1.year - d2.year) * 12 + d1.month - d2.month


def read_password_file(filepath, password, sheet=None, header=0, dtype=None, encoding='utf-8', skiprows=None, usecols=None):
    """Processes Password protected excel into DataFrame
    If the file contains multiple sheets, currently the user needs to call this function multiple times with explicit
    sheet names.
    """
    f = NamedTemporaryFile(delete=False, suffix='.csv')
    f.close()
    os.unlink(f.name)  # Not deleting will result in a "File already exists" warning
    xlwb = xlApp.Workbooks.Open(Filename=filepath, UpdateLinks=False, ReadOnly=True, Format=None, Password=password)
    xlCSVWindows = 0x17  # CSV file format, from enum XlFileFormat

    if sheet:
        logStatus("Reading", sheet)
        ws = xlwb.Worksheets(sheet)
        ws.SaveAs(Filename=f.name, FileFormat=xlCSVWindows)
    else:
        xlwb.SaveAs(Filename=f.name, FileFormat=xlCSVWindows)  # Save the workbook as CSV

    outputdf = pd.read_csv(f.name, header=header, dtype=dtype, encoding=encoding, skiprows=skiprows, usecols=usecols)  # Read that CSV from Pandas
    xlwb.Close(False)
    return outputdf


def send_emails(email_to=None, email_cc=None, email_subject=None, email_body=None, email_attachment=None):
    """Send emails automatically (mostly warning messages)"""
    outlook = client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    if not email_to and not email_cc:
        logStatus("There must be at least one email address")
    if email_to:
        mail.To = email_to
    if email_cc:
        mail.CC = email_cc
    if email_subject:
        mail.Subject = email_subject
    if email_body:
        mail.Body = email_body

    # To attach a file to the email (optional):
    if email_attachment:
        attachment = email_attachment
        mail.Attachments.Add(attachment)

    mail.Send()


def dtype_category(category_col):
    """

    :param category_col:
    :return: dtype: a dictionary {col: 'category'}
    """
    dtype = {}
    for c in category_col:
        dtype[c] = 'category'
    return dtype


def fillna_category_col(df, cols, value):
    """

    :param df:
    :param cols:
    :param value: fillna value. usually it's TBD or NA
    :return:
    """
    df[cols] = df[cols].astype('category')
    for col in cols:
        if value not in df[col].cat.categories:
            df[col] = df[col].cat.add_categories([value])
    df[cols] = df[cols].fillna(value)
    return df
