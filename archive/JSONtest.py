# -*- coding: utf-8 -*-
"""
Created on Thu Mar 12 16:59:44 2020

@author: user
"""
import json
filename = "test_unicode_in.json"

data_list = ["adadad"] # [["@",""],[":",""],["'",""],['"',''],["{",""],["}",""],["±","+\- "],["Ø","диам. "],["®",""],["™",""],["мама","штепсель-розетка"],["папа","штепсель-вилка"],["Монтаж THT","Монтаж в отверстия печатной платы"],["Монтаж PCB","Монтаж на печатную плату"],["Монтаж SMD","Монтаж на поверхность печатной платы"],["DC","постоянного тока"],["AC","переменного тока"]]

data = json.load(open(filename))
print(data)

with open(filename,'w') as file:
    json.dump(data_list, file, indent=2, ensure_ascii=False)