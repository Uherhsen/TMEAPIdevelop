# -*- coding: utf-8 -*-
"""
Created on Tue Jan 14 12:02:49 2020

Получение данных о электронных компонентах с сайта TME через API
"""
import win32com.client
from TME_Python_API import product_import_tme

Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(u"d:\\usr\\documents\\Desktop\\TMEAPIdevelop\\APP\\productdata.xlsx") #путь к файлу
sheet = wb.ActiveSheet

i = 1
sumcells = ""
while sheet.Cells(i,1).value != None:
    sumcells += sheet.Cells(i,1).value
    i+=1
    if len(sumcells) >= 20:
        #print('длина артикулов: ',len(sumcells),'\nномер строки экселя: ',i-1, sumcells)
        break   
cord = 'A'+str(1)+':A'+str((i-1)) 


work_articles_list = [r[0].value for r in sheet.Range(cord)]

product_key = 'SymbolList['
SymbolList = {}
for j in range(len(work_articles_list)):
    SymbolList[product_key+str(j)+']'] = work_articles_list[j]

CountryLanguageParams={'Country' : 'RU','Language' : 'RU',}
params={**CountryLanguageParams,**SymbolList}
token = 'ac434c181917ed4e51c49a2027bfd040e9f2da0054be7'
app_secret = '0b748f6e5d340d693703'
action = 'Products/GetProducts' # request method

all_data = product_import_tme(token, app_secret, action, params) 

for j in range(len(work_articles_list)): 
    print('Артикул: '+SymbolList['SymbolList['+str(j)+']'],'\nкраткое описание: ', all_data['Data']['ProductList'][j]['Description'],'\nВес: ', all_data['Data']['ProductList'][j]['Weight']*0.001)
    
wb.Close()
#Excel.Quit()