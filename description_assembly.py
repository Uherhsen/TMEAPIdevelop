# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 16:00:40 2020

Программа для сборки описания по информации из экселя и из шаблонов JSON БД (Работает уже с заполненным экселем)
"""
from get_product_name import product_type
import GetDataTME_with_openpyxl as Data
import openpyxl,json

path = "productdata.xlsx"
filename = 'database.json'
wb = openpyxl.load_workbook(path)#путь к файлу
sheet = wb.active
list_range = Data.number_of_articles(path)
descr_list = Data.articles_list(list_range,"G", path) # список параметров

'''Проверка и обновление типов продуктов в базе JSON - database.json'''

# Список типов из экселя,- слова до точки-запятой в колонке description
categ_list = [product_type(d) for d in descr_list]

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
    первый словарь - это множество типов товаров полученное отделением типа от {описаний} в качестве ключей, по которым получаем значение-ссылку на ключ второго словаря с описаниями, 
    второй словарь - по ключу отдает описание в виде списка [0-Начало описания,1-конец описания,[[заменяемое,заменитель],[..., ...]]].
описание составляется конкатенацией: нулевой элемент словаря + параметры из экселя + первый элемент словаря, при этом происходят общие замены (некоторые символы не допускаются в описании)
затем происходят частные замены, которые указываются в индивидульном списке в описании из БД
Если во втором словаре при обращении по ключу первый (с нулевым индексом) элемент словаря == "MY_FUNСTION", то запускается альтернативный сценарий, который распаковывает строку-функцию с индексом [1] в этом же словаре'''

# Список для общих замен
replace_list = [["@",""],[":",""],["'",""],['"',''],["{",""],["}",""],["±","+\- "],["Ø","диам. "],["®",""],["™",""], ["мама","штепсель-розетка"],["папа","штепсель-вилка"],["Монтаж THT","Монтаж в отверстия печатной платы"],["Монтаж PCB","Монтаж на печатную плату"],["DC","постоянного тока"],["AC","переменного тока"]]

def just_parameters(i):
    print("\nВ ячейку сохранены только параметры, без описания: \n\n"+ prod_param+".\n")
    sheet["I"+str(i)]= 'Параметры без описания:\n '+prod_param+'.'

# Заготовка
def desc_assembly():
    pass
    
for i in range(1, list_range+1):
    print(i)
    try:
        # Отделяем категорию от дескрипшина
        name = product_type(sheet['G'+str(i)].value)
        prod_param = sheet['G'+str(i)].value
        
        # цикл общих замен
        for a,b in replace_list:
            prod_param = prod_param.replace(a,b)
        try:
            # Переменная-ссылка-ключ к словарю с описаниями
            link_key = data[0][name]
            # Проверка наличия флага "MY_FUNСTION" для включения иных сценариев
            if data[1][link_key][0] == "MY_FUNСTION":
                # Включение функции написанной в базе
                eval(data[1][link_key][1])
            else:   
                try:
                    my_descr = data[1][link_key][0]+prod_param+data[1][link_key][1]
                except TypeError:
                    print('Шаблон-описание пуст')
                    continue
                except IndexError:
                    print('Ошибка в шаблоне-описании')
                    continue
                try:
                    for x,y in data[1][link_key][2]:
                        my_descr = my_descr.replace(x,y)
                    print('С доп. заменой:\n'+ my_descr)
                    sheet["I"+str(i)]= 'С доп. заменой:\n'+ my_descr
                except IndexError:
                    print(my_descr)
                    sheet["I"+str(i)]= my_descr
        except KeyError:
            print("Не указан шаблон-описание")
            continue
    except AttributeError:
        print('Артикула нет на ТМЕ')
        continue
    
wb.save(path)

#print(eval(prod_param)['Монтаж']) # Конструкция для превращения строки в словарь, опасно!