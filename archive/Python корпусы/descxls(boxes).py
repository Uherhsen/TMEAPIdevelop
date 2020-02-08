# -*- coding: utf-8 -*-
#https://habr.com/ru/post/232291/ 
#'D:/Desc.xlsx'
import win32com.client, re
#numpos = input('Сколько позиций?\n')
#numposint = int(numpos)
cord = 'A1:A7329'#+str(numposint) #НОМЕР СТОЛБЦА И СТРОКИ интервал
#print(numpos,cord)
Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(u"D:\\BoxesDesc.xlsx") #путь к файлу
sheet = wb.ActiveSheet

#получаем значение первой ячейки (строка,столбец)
#val = sheet.Cells(2,1).value
#получаем значения цепочки A1:A2
vals = [r[0].value for r in sheet.Range(cord)]
#print(re.findall(r'\w\d+',val))
#val2 = re.findall(r'\w\d+',val)
#записываем значение в определенную ячейку
#sheet.Cells(2,12).value = val2[1]
#записываем последовательность
mydesc0 = 'Корпусы '
mydesc1 = 'изготовлены из '
mydesc2 = '. Обеспечивает защиту от проникновения пыли и влаги, габаритные размеры составляют '
mydesc2_2 = '. Обеспечивает защиту от проникновения пыли и влаги'
mydesc3= '. На внутренней поверхности корпуса отлиты стойки и направляющие для размещения печатных плат. Предназначены для использования в качестве корпусов для размещения электротехнических сборок.'

body_types = {'универсальный':'универсальные, укомплектованные винтами, ','для USB':'с отверстием для USB разъёма, ','специализированный':'универсальные, ', 'для модульных устройств':'для монтажа модульных устройств, ', 
              'встраиваемый':'универсальные, встраиваемые, ', 'для мультимедийных ПК':'', 'для пультов':'для размещения платы дистанционного управления, ',  
              'для устройств с дисплеем':'для устройств с дисплеем ,','для сигнализации':'универсальные, для настенного монтажа, ', 'настенный':'универсальные, для настенного монтажа ,',
              'под заливку':'универсальные, ','для блоков питания':'для размещения блока питания, ','for power supplies':'для размещения блока питания,','на DIN-рейку':'для монтажа на din-рейку, ',
              'с панелью':'универсальные, с панелью для размещения элементов управления, ', 'приборный':'универсальные, с панелью для размещения элементов управления, ',
              'стандарта 19':'под установку в слот стойки стандарта 19-дюймов, ','противопожарный':'универсальные, ','на панель':'универсальные, для монтажа на панель, ','экранирующая':'универсальные, экранирующие, '}

body_keys= ['универсальный', 'для USB', 'специализированный', 'для модульных устройств', 'встраиваемый', 'для мультимедийных ПК',
            'для пультов', 'для устройств с дисплеем', 'для сигнализации', 'настенный', 'под заливку', 'для блоков питания',
            'for power supplies', 'на DIN-рейку', 'экранирующая', 'с панелью', 'приборный', 'стандарта 19', 'противопожарный','на панель','экранирующая']

unknown_material = 0

i = 1
size=[]
text=''
y=0
for rec in vals:
    rec = str(rec)
    print(rec)
    for j in body_keys:
        y=rec.find(j)
        if y > -1:
            sheet.Cells(i,9).value = body_types[j]
            break
        else:
            sheet.Cells(i,9).value = 'универсальные, '
    size = re.findall(r'(\d*\S*\d+)мм',rec)
    #print(size)
    if len(size)>2:
        text = (size[0]+' x '+size[1]+' x '+size[2]+' мм')
        sheet.Cells(i,11).value = text#size[0],size[1],size[2]#(size[0],'мм',size[1],'мм',size[2],'мм') #str(rec)+'ёёбёбёбё'
        i = i + 1
    else:
        text = 'Нет размеров.'
        sheet.Cells(i,11).value = text#size[0],size[1],size[2]#(size[0],'мм',size[1],'мм',size[2],'мм') #str(rec)+'ёёбёбёбё'
        i = i + 1
                   
i = 1
mat = ''
code = 0
for rec2 in vals:
    rec2 = str(rec2)
    word1 = rec2.find('ABS')
    word2 = rec2.find('поликарбонат')
    word3 = rec2.find('алюминий')
    word4 = rec2.find('ALU')
    word5 = rec2.find('полиэфир')
    word6 = rec2.find('полистирен')
    word7 = rec2.find('сталь')
    word8 = rec2.find('полипропилен')
    word9 = rec2.find('полиамид')
    if word1 > -1:
        mat = 'ABS-пластмассы'
        code = 3926909709
    elif word2 > -1:
        mat = 'поликарбоната'
        code = 3926909709
    elif word3 > -1:
        mat = 'алюминия'
        code = 7616999008
    elif word4 > -1:
        mat = 'алюминия'
        code = 7616999008
    elif word5 > -1:
        mat = 'полиэфира'
        code = 3926909709
    elif word6 > -1:
        mat = 'полистирена'
        code = 3926909709
    elif word7 > -1:
        mat = 'легированной стали, комбинированным способом'
        code = 7326909807
    elif word8 > -1:
        mat = 'полипропилена'
        code = 3926909709  
    elif word9 > -1:
        mat = 'полиамида'
        code = 3926909709        
    else:
        mat = 'поликарбоната'
        code = 3926909709
        sheet.Cells(i,15).value = '!!!ПРОВЕРИТЬ МАТЕРИАЛ!!!!!'
        unknown_material += 1
    sheet.Cells(i,12).value = mat
    sheet.Cells(i,10).value = code
    if str(sheet.Cells(i,11)) != 'Нет размеров.':
        #print(sheet.Cells(i,11))
        sheet.Cells(i,14).value = mydesc0+str(sheet.Cells(i,9).value)+mydesc1+mat+mydesc2+str(sheet.Cells(i,11).value)+mydesc3
    else:
        sheet.Cells(i,14).value = mydesc0+str(sheet.Cells(i,9).value)+mydesc1+mat+mydesc2_2+mydesc3
    i = i + 1 

print(unknown_material)      
#сохраняем рабочую книгу
wb.Save()
#закрываем ее
#wb.Close()
#закрываем COM объект
#Excel.Quit()