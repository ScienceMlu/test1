# _*_ coding=utf-8 _*_
import os
import openpyxl
from openpyxl.styles import Font, colors, Alignment

wb = openpyxl.load_workbook('1.xlsx')

sheet = wb['Sheet1']

print(sheet['A1'].fill.fgColor.rgb)

print(sheet['C1'].fill.fgColor.rgb)
print(sheet['E1'].fill.fgColor.rgb)
print(sheet['G1'].fill.fgColor.rgb)
print(sheet['I1'].fill.fgColor.rgb)
print(sheet['B1'].value)


