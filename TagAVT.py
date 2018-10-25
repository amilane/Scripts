# -*- coding: utf-8 -*-

import clr
import System
from System import Array
import math

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")
import string
from operator import itemgetter
from itertools import groupby

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import *
from System.Windows.Forms import *

NONELIST = [None, '', ' ']
CODE = doc.ProjectInformation.LookupParameter('CODE').AsString()

def groupByParameter(elems, param):
  sElems = sorted(elems, key = lambda e: e.LookupParameter(param).AsString())
  gElems = [[x for x in g] for k,g in groupby(sElems, lambda e: e.LookupParameter(param).AsString())]
  return gElems

excel = Excel.ApplicationClass()
excel.Visible = False
path = doc.ProjectInformation.LookupParameter('Path to BD').AsString()
workbooks = excel.Workbooks
workbook = workbooks.Open(path)
worksheet = workbook.Worksheets[1]
"""
uRange = worksheet.UsedRange
excelEquipTags = list(uRange.Columns(1).Value2)
excelCableTags = list(uRange.Columns(26).Value2)
excelSignalTags = list(uRange.Columns(32).Value2)
# задействованное кол-во строк
uRows = len(excelEquipTags)
"""

elems = FilteredElementCollector(doc).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_DetailComponents).WhereElementIsNotElementType().ToElements()
#Элементы с заполненным GroupTag и Описанием
elemsG = filter(lambda e: e.LookupParameter('GroupTag').AsString() not in NONELIST\
and e.Symbol.LookupParameter('Element Type').AsString() not in NONELIST, elems)

elemsGroupByTag = groupByParameter(elemsG, 'GroupTag')
lol = []
for group in elemsGroupByTag:

  uRange = worksheet.UsedRange
  excelEquipTags = list(uRange.Columns(1).Value2)
  excelCableTags = list(uRange.Columns(26).Value2)
  excelSignalTags = list(uRange.Columns(32).Value2)

  equip = []
  signal = []
  for elem in group:
    elemType = elem.Symbol.LookupParameter('Element Type').AsString()
    if 'EQ' in elemType:
      equip.append(elem)
    if 'S' in elemType:
      signal.append(elem)
  if equip:
    equipTags = ['-'.join([CODE, e.LookupParameter('partition').AsString(), e.LookupParameter('Signal_designation').AsString(), e.LookupParameter('EQ_TAG').AsString()]) for e in equip]
    lol.append(equipTags)

  #a = Array.CreateInstance(str, len(equipTags))
  a = Array[str](equipTags)
  xlEquipRange = worksheet.Range["A"+str(len(excelEquipTags)+1) + ":A"+str(len(excelEquipTags)+len(equipTags))]
  #for n, i in enumerate(equipTags):
  #  a[n] = i
  xlEquipRange.Value2 = a


MessageBox.Show(str(lol), "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)
workbook.Save
workbook.Close(1)
workbooks.Close()

excel.Quit()

Marshal.ReleaseComObject(uRange) 
Marshal.ReleaseComObject(worksheet)
Marshal.ReleaseComObject(workbook)
Marshal.ReleaseComObject(workbooks)
Marshal.ReleaseComObject(excel)


#MessageBox.Show(str(dir(worksheet)), "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)

