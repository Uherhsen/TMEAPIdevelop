# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 16:00:40 2020

Программа для сборки описания по информации из экселя и из шаблонов JSON БД (Работает уже с заполненным экселем)
"""
from get_product_name import product_name
import GetDataTME_with_openpyxl as Data
import openpyxl,json

path = "X:\\PythonProjects\\TMEAPIdevelop_v9\\APP\\productdata.xlsx"
filename = 'database.json'

wb = openpyxl.load_workbook(path)#путь к файлу
sheet = wb.active
list_range = Data.number_of_articles(path)
descr_list = Data.articles_list(list_range,"C", path) # список параметров

'''Проверка и обновление типов продуктов в базе JSON - database.json'''

# Список типов из экселя,- слова до точки-запятой в колонке description
categ_list = [product_name(d) for d in descr_list]

# Получаем словарь значений из базы
try:
    data = json.load(open('database.json')) # FileNotFoundError:
except FileNotFoundError:
    print('Файл базы database.json отсутствует, или неверно указанно имя файла')

# Получаем список ключей из словаря базы json (ключи это типы продукции, - первые слова дескрипшина только уже помещенные в базу )
data_keys = list(data[0].keys()) 
    
'''подгружает новые элементы, присваивает значения None (null стандарт джейсона)'''
for e in categ_list:
    # Добавляем данные в словарь, избегая повторений
    if e not in data_keys: data[0][e] = None

# Звгружаем обновленный словарь в базу json
with open(filename,'w') as file:
    json.dump(data, file, indent=2, ensure_ascii=False)

''' Сборка описания. 
БД database.json содержит список из двух словарей: 
    первый словарь - это множество типов товаров полученное отделением типа от дескрипшина в качестве ключей, по которым получаем значение-ссылку на ключ второго словаря с описаниями, 
    второй словарь - по ключу отдает описание в виде списка [0-Начало описания,1-конец описания,[[заменяемое,заменитель],[..., ...]]].
описание составляется конкатенацией: нулевой элемент словаря + параметры из экселя + первый элемент словаря, при этом происходят общие замены (некоторые символы не допускаются в описании)
затем происходят частные замены, которые указываются в индивидульном списке в описании из БД'''

# Список для общих замен
replace_list = [["'",""],["{",""],["}",""],[":",""],["±","+\- "],["Ø","диам. "],["®",""],["™",""]] 

def desc_assembly():
    pass
    
for i in range(1,list_range+1):
    print(i)
    try:
        # Отделяем категорию от дескрипшина
        name = product_name(sheet['C'+str(i)].value) 
        prod_param = sheet['G'+str(i)].value
        # цикл общих замен
        for a,b in replace_list:
            prod_param = prod_param.replace(a,b)
        try:
            # Переменная-ссылка-ключ к словарю с описаниями
            link_key = data[0][name]
            try:
                my_descr = data[1][link_key][0]+prod_param+data[1][link_key][1]
            except TypeError:
                print('Шаблон-описание пуст')
                continue
            except IndexError:
                print('Ошибка в шаблоне-описании')
                continue
            try:
                for i,j in data[1][link_key][2]:
                    my_descr = my_descr.replace(i,j)
                print('С доп. заменой:\n'+ my_descr)
            except IndexError:
                print(my_descr)
        except KeyError:
            print("Не указан шаблон-описание")
            continue
    except AttributeError:
        print('Артикула нет на ТМЕ')
        continue
        
wb.save(path)

#print(eval(prod_param)['Монтаж']) # Конструкция для превращения строки в словарь, опасно!