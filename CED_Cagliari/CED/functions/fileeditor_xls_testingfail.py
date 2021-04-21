import xml.etree.ElementTree as ET
import pandas as pd
import xlwt
import os


paflist = dict()
paflist["month"] = dict()
#paflist["month"]["service"] = "SOS PC"
paflist["month"]["progress"] = "22/10"
paflist["month"]["consuntivato_pezzi"] = 12852
paflist["month"]["consuntivato_euro"] = 135122.22
paflist["month"]["preventivato_pezzi"] = 14568

pathstrings = r".././xls_templates/xl/sharedStrings.xml"
pathsheet1 = r".././xls_templates/xl/worksheets/sheet1.xml"
ET.register_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
treestring = ET.parse(pathstrings)
rootstr = treestring.getroot()
x = 0
z = 0
sharderstr = list()
unique = int(rootstr.attrib.get("uniqueCount"))
total = int(rootstr.attrib.get("count"))

listkey = list(paflist["month"])

for si in rootstr:
    x +=1
    if x == (unique):
        for i in range(len(listkey)):
            if isinstance(paflist["month"][listkey[i]], str):
                si = ET.SubElement(rootstr, 'si')
                t = ET.SubElement(si, 't')
                t.text = paflist["month"][listkey[i]]
                z += 1
                sharderstr.append(str(x + z))

rootstr.set("uniqueCount", str(unique+z))

treesh = ET.parse(pathsheet1)
rootsh = treesh.getroot()
x = 0
cells = dict()
cells["type"] = ["B10", "B18", "B26"]
cells["consuntivo"] = ["B26", "I18"]
writecells = 0

for sheet in rootsh:
    x += 1
    if x == 6:
        for row in sheet:
            for cell in row:
                if cell.attrib.get("r") == "B3":
                    cell.set("t", "s")
                    cell.set("s", "181")
                    v = ET.SubElement(cell, 'v')
                    v.text = sharderstr[0]
                    writecells +=1


rootstr.set("count", str(total+writecells))
treestring.write(r".././prova/xl/sharedStrings.xml", encoding = 'utf-8', xml_declaration=True)

#ET.register_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
ET.register_namespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

treesh.write(r".././prova/xl/worksheets/sheet1.xml", encoding = 'utf-8', xml_declaration=True)
