# -*- coding: utf-8 -*-
"""
Created on Sun Feb  2 10:39:57 2020

Учимся использовать openpyxl для загрузки и выгрузки экселя

Sheet1
productdata.xlsx
X:\\PythonProjects\\excel\\
"""
import openpyxl
wb = openpyxl.load_workbook('X:\\PythonProjects\\excel\\productdata.xlsx')


sheet = wb.active
# Получение значения
val = sheet['A1'].value # значение из ячейки

# Получение диапазона значений в список
vals = [v[0].value for v in sheet['A1':'A49']]

# Запись в файл из списка vals
i = 1
for rec in vals:
    print(rec)
    sheet.cell(column=2, row = i, value=rec )
    i += 1

wb.save('X:\\PythonProjects\\excel\\productdata.xlsx')