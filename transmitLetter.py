import openpyxl
import os
from PyPDF2 import PdfFileReader

number = '0055-P2-ARM-CPC-TRM-00008'

path = 'D:\\' + number + '\\Documents'

pathToVDR = 'D:\\0055-P2-ARM-CPC-TRM-00008\\0055-CPC-ARM-4.1.0.00.CPC-OVK-TP-0001_A1_ER.xlsx'
wbVDR = openpyxl.load_workbook(pathToVDR, data_only=True)
vdrSheet = wbVDR.get_sheet_by_name('VDR')
vdrRows = range(12, 291)

# 29-fileName, 16-docClass, 25-disc, 27-docType, 28-docNo, 30-nameDocEn, 31-nameDocRu, 34-issue, 35-Rev, 36-order, 41-date
vdrDict = {vdrSheet.cell(row = row, column = 29).value:\
            [vdrSheet.cell(row = row, column = 16).value,\
            vdrSheet.cell(row = row, column = 25).value,\
            vdrSheet.cell(row = row, column = 27).value,\
            vdrSheet.cell(row = row, column = 28).value,\
            vdrSheet.cell(row = row, column = 30).value,\
            vdrSheet.cell(row = row, column = 31).value,\
            vdrSheet.cell(row = row, column = 34).value,\
            vdrSheet.cell(row = row, column = 35).value,\
            vdrSheet.cell(row = row, column = 36).value,\
            vdrSheet.cell(row = row, column = 41).value\
            ] for row in vdrRows}

os.chdir(path)
folders = os.listdir()


transData = []
transData_CSV = []

for folder in folders:
  pathToFolder = path + '\\'+ folder
  os.chdir(pathToFolder)
  files = os.listdir()
  pdfs = list(filter(lambda f: '.pdf' in f, files))
  for pd in pdfs:
    language = pd.split('_')[-1].split('.')[0]
    packageCode = pd.split('-')[4]
    pathToPdf = pathToFolder + '\\' + pd
    pdf = PdfFileReader(open(pathToPdf,'rb'))
    numberOfSheets = pdf.getNumPages()
    for n in range(numberOfSheets):
      page = pdf.getPage(n)
      box = page.mediaBox
      upperRight = box.upperRight
      w = upperRight[0]
      h = upperRight[1]
    sheetFormat = 'sheetFormat'
    try:
      vdrRow = vdrDict[pd]
      tdRow = [vdrRow[8],\
              None, \
              None, \
              vdrRow[3], \
              language, \
              vdrRow[5], \
              vdrRow[4], \
              vdrRow[1], \
              packageCode, \
              vdrRow[6], \
              vdrRow[9], \
              vdrRow[0], \
              vdrRow[7], \
              vdrRow[2], \
              numberOfSheets, \
              sheetFormat, \
              'pdf', \
              pd]
    except:
      tdRow = [None,\
              None, \
              None, \
              None, \
              language, \
              None, \
              None, \
              None, \
              packageCode, \
              None, \
              None, \
              None, \
              None, \
              None, \
              numberOfSheets, \
              sheetFormat, \
              'pdf', \
              pd]
    transData.append(tdRow)



print(w, h)
