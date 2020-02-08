# -*- coding: utf-8 -*-
"""
Created on Mon Feb  3 16:02:01 2020

Программа для заполения базы данных JSON из экселя, наименованием товара
получает наименование товара из дескрипшина (ТО, что до точки с запятой)
"""
import GetDataTME_with_openpyxl as Data
from input_productname_json import json_load
from get_product_name import product_name
import openpyxl

path = "D:\\python_programming\\TMEAPIdevelop_v7\\APP\\productdata.xlsx"
filename = 'typebase.json'
wb = openpyxl.load_workbook(path)#путь к файлу
sheet = wb.active
rng=Data.number_of_articles(path)
for j in range(rng):
    #print(product_name(sheet['C'+str(j+1)].value))
    json_load(product_name(sheet['C'+str(j+1)].value),filename)
       
wb.save(path)
print('\nГотово')