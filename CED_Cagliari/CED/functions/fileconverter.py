import xml.etree.ElementTree as ET
import pandas as pd
import xlwt
import os

# from xml to excel
def convert_xml(path):
    pathorg = r"./excel/origin.xml"
    try:
        os.rename(path, pathorg)
    except FileNotFoundError:
        None
    tree = ET.parse(pathorg)
    root = tree.getroot()
    x = 0
    df = dict()

    wbook = xlwt.Workbook()
    wsheet = wbook.add_sheet("sheet")

    for ws in root:
        for table in ws:
            for row in table:
                y = 0
                x = x + 1
                for cell in row:
                    y = y + 1
                    for data in cell:
                        wsheet.write(x-27, y-1, data.text)

    wbook.save(r"./excel/converted.xls")
    df = pd.read_excel(r"./excel/converted.xls", dtype="object")
    df.to_csv(r"./excel/converted.csv", sep=";", index=False)
    os.remove(pathorg)


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
