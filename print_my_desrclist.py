# -*- coding: utf-8 -*-
"""
Created on Thu Feb 20 16:21:37 2020

@author: user
"""

import json

def my_descrlist(jsonname):
    '''Проверяет наличие ДБ, гружает '''
    try:
        data = json.load(open(jsonname))
        return list(data[1].keys())
    except:
        print("Отсутствует файл ДБ "+jsonname)
        
if __name__ == "__main__":
    filename = 'database.json'
    print('Отсортированный:\n\n',sorted(my_descrlist(filename)),'\n\nНе отсортированный:\n\n',my_descrlist(filename))