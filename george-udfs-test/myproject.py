import xlwings as xw
import pandas as pd
import re

# Excel 
# https://stackoverflow.com/questions/11243423/getting-email-address-from-a-cell-in-excel
# Python https://www.tutorialspoint.com/python_text_processing/python_extract_emails_from_text.htm  




def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"


@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("myproject.xlsm").set_mock_caller()
    main()

########


@xw.func
def py_find_emails(str):
    adds = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", str)
    return adds



@xw.func
@xw.arg("df", pd.DataFrame, index=False, header=True)
def py_describe(df):
    return df.describe()


@xw.func
@xw.arg("df", pd.DataFrame, index=False, header=True)
def py_corr(df):
    return df.corr()
