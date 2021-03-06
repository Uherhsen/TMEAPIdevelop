# -*- coding: utf-8 -*-
"""
Created on Mon Mar 16 14:25:12 2020

проверка if not in {:}
"""
import json 
data = json.load(open('database для теста разъёмов.json'))

# 1 Словарь поступающий из экселя
prms={'Разъем': 'штыревая планка', 'Тип разъема': 'штыревой', 'Серия разъема': 'AMPMODU MOD II',
      'Вид разъемов': '"папа"', 'Кол-во выводов': '40', 'Пространственная ориентация': 'прямой', 
      'Шаг контактов': '2,54мм', 'Электрический монтаж': 'THT', 'Конфигурация выводов разъема': '1x40',
      'Размеры выводов': 'дл. 3,18мм', 'Высота': '8,08мм'}#,'Покрытие контакта': 'луженые'} 
# Имитация переменнолй со значением ключа
link_key = data[0]["Тип разъема провод - плата"]

# 2 Словарь значений по-умолчанию (проверяет и заменяет если не нашел ключи)
#prmsDflt={'Электрический монтаж':'на кабель','Кол-во выводов':'10','Шаг контактов':'2.54 мм','Конфигурация выводов разъема':'1х5','Покрытие контакта':'золота'}
# 3 список замен
#replaceL = (['THT','в отверстия печатных плат'],['пайка','на проводники'], ['IDC','на проводники'],['луженые','олова'])
# шаблон описание
descTemplate = "Штепсельные разъемы для монтажа {0[Электрический монтаж]}, {0[Кол-во выводов]} контактов расположены с шагом {0[Шаг контактов]}, c конфигурацией выводов {0[Конфигурация выводов разъема]}. Номинальное напряжение до 50 В, номинальный ток 3 А, контакты изготовлены из медного сплава с гальваническим покрытием из {0[Покрытие контакта]}, материал корпуса полиамид, рабочие температуры от -40 до 105°С, предназначены для радиоэлектронного оборудования общепромышленного назначения."
           
#key_prmsDflt =[*prmsDflt] 
def textTemplate(prms,link_key):   
    #prms = eval(sheet['G'+str(i)].value)
    #print(data[1][link_key][1])
    descTemplate = data[1][link_key][2]
    prmsDflt = data[1][link_key][3]
    
    for j in [*prmsDflt]:
        if j not in prms:
            prms[j]= prmsDflt[j]
    d=descTemplate.format(prms)
    try:
        replaceL = data[1][link_key][4]
        for a,b in replaceL:
            d = d.replace(a,b)
        return d
    except IndexError:
        print('нет списка замен')
        return d    
    
# Реализация сложения словарей для замен
a=['one','two','three','four']
b=[1,2,3]
ab=a+b
#ab=[*b,*a]
print(ab)

# получить имя функции:
def getMyName():
    print('меня зовут '+getMyName.__name__)
getMyName()

print(eval(data[1]["connectors_func"][1]))