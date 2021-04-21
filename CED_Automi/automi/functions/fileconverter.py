import xml.etree.ElementTree as ET
import pandas as pd
import xlwt
import os


# from html to excel
def renandconv_HTML(path):
    pathorg = r"./excel/ExportFuoriSTD.html"
    try:
        os.rename(path, pathorg)
    except FileNotFoundError:
        None

    df_list = pd.read_html(pathorg)
    for i, df in enumerate(df_list):
        df.to_excel(r'./excel/import_html.xls'.format(i), index=False, header=False)
