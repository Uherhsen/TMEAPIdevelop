# -*- coding: utf-8 -*-
"""
Created on Sun Feb  2 09:24:17 2020
Создает json бд, добавляет в нее элмементы
"""
import json

def json_load(datain,jsonname):
    '''Проверяет наличие ДБ, гружает новые элементы'''
    try:
        data = json.load(open(jsonname))
    except:
        data = []
    # Добавляем избегая повторений
    if datain not in data: data.append(datain)
    # Отсортировать
    sorted(data)
    
    with open(jsonname,'w') as file:
        json.dump(data, file, indent=2, ensure_ascii=False)


if __name__ == "__main__":
    filename = 'typebase.json'
    inpt='y'
    
    print('Записываем элементы в простой массив дт json\n')
    
    while inpt !='n':
        datainput = input('Добавить элемент:\n')
        
        json_load(datainput)
        
        inpt = input('Продолжить ? n - нет \n')