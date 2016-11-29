# Please use
#	python -m pdb DemoScript.py
# and press 'n' to execute the script line by line. To exit, type 'q'

import shutil
shutil.copyfile('blank.xlsx', 'demo.xlsx')

import os
os.startfile('demo.xlsx')

import excel
import exceldemolib

exceldemolib.ExcelDemoLib.initDesktopContext()
context = excel.RequestContext()

exceldemolib.ExcelDemoLib.populateData(context)

exceldemolib.ExcelDemoLib.analyzeData(context)

range = context.workbook.worksheets.getItem('Sheet1').getRange('B2:C3')
context.load(range)
context.sync()
print(range.values)
print(range.text)

imageBase64 = exceldemolib.ExcelDemoLib.getChartImage(context)


shutil.copyfile('blank.docx', 'demo.docx')
os.startfile('demo.docx')

import word
import worddemolib
worddemolib.WordDemoLib.initDesktopContext()
context = word.RequestContext()
worddemolib.WordDemoLib.insertPictureAtEnd(context, imageBase64)


