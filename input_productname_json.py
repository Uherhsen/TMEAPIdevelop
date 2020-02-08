# -*- coding: utf-8 -*-
"""
Created on Sun Feb  2 09:24:17 2020
Создает json бд, добавляет в нее элмементы
"""
import json

inpt='y'
print('Записываем элементы в простой массив дт json\n')
while inpt !='n':
    datain = input('Добавить элемент:\n')
    
    try:
        data = json.load(open('database.json'))
    except:
        data = []
    data.append(datain)
    
    with open('database.json','w') as file:
        json.dump(data, file, indent=2, ensure_ascii=False)
    
    inpt = input('Продолжить ? n - нет \n')
