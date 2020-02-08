# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 16:00:40 2020

Программа для сборки описания по информации из экселя и из шаблонов JSON БД (Работает уже с заполненным экселем)
"""
from get_product_name import product_name
import GetDataTME_with_openpyxl as Data
import openpyxl,json

path = "D:\\python_programming\\TMEAPIdevelop_v7\\APP\\productdata.xlsx"
filename = 'database.json'

#def json_load_none():
#    '''Создаем словь с значениеми None, ключами из списка typebase.json'''
#    '''ВНИМАНИЕ!!! ЭТА ФУНКУЦИЯ УНИЧТОЖИТ ЗНАЧЕНИЯ!'''
#    data = json.load(open('typebase.json'))
#    with open('database.json','w') as file:
#            json.dump(dict.fromkeys(data), file, indent=2, ensure_ascii=False)

wb = openpyxl.load_workbook(path)#путь к файлу
sheet = wb.active
list_range = Data.number_of_articles(path)
descr_list = Data.articles_list(list_range,"C",path) # список параметров

categ_list = []
for d in descr_list:
    d = product_name(d)
    categ_list.append(d) # Список типов из экселя
    
# Получаем словарь значений из базы
try:
    data = json.load(open('database.json')) # FileNotFoundError:
except FileNotFoundError:
    print('Файл базы database.json отсутствует, или неверно указанно имя файла')

# Получаем список ключей из словаря базы json (ключи это типы)
data_keys = list(data.keys()) 
    
'''подгружает новые элементы, присваивает значения None'''

for e in categ_list:
    # Добавляем избегая повторений
    if e not in data_keys: data[e]= None
    # Отсортировать
    sorted(data)
    
    with open(filename,'w') as file:
        json.dump(data, file, indent=2, ensure_ascii=False) 

for i in range(1,list_range+1):
    print(i)
    # Отделяем категорию от дескрипшина
    try:
        name = product_name(sheet['C'+str(i)].value) 
        prod_param = sheet['G'+str(i)].value
    
        prod_param = prod_param.replace("'","")
        prod_param = prod_param.replace("{","")
        prod_param = prod_param.replace("}","")
        prod_param = prod_param.replace(":","")
        try:
            print(data[name][0]+prod_param+data[name][1])
        except TypeError:
            print('Нет шаблона описания')
            continue
    except AttributeError:
        print('Артикула нет на ТМЕ')
        continue
    
        
wb.save(path)

#print(eval(prod_param)['Монтаж']) # Конструкция для превращения строки в словарь, опасно!

# Может использовать MAP что бы получить список типов и 
#обработать их функцией преобразующей в описание