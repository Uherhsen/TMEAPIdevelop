# -*- coding: utf-8 -*-
"""
Created on Tue Jan 14 12:02:49 2020

Получение данных о электронных компонентах с сайта TME через API
"""
import win32com.client,time
from TME_Python_API import product_import_tme

Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(u"d:\\usr\\documents\\Desktop\\TMEAPIdevelop\\APP\\productdata.xlsx") #путь к файлу
sheet = wb.ActiveSheet

def cord (a1,a2):
    return 'A'+str(a1)+':A'+str(a2)
    
def sheet_range(i):
    column_range=[]
    column_range.append(i)
    sumcells = ""
    while sheet.Cells(i,1).value != None:
        sumcells += sheet.Cells(i,1).value
        i+=1
        if len(sumcells) >= 20:
            column_range.append(i-1)
            break
    print('sumcells',sumcells,'len', len(sumcells))
    return column_range

def get_data(i):
    new_cord = sheet_range(i)
    print(new_cord)
    work_articles_list = [r[0].value for r in sheet.Range(cord(new_cord[0],new_cord[1]))]
    print(work_articles_list)
    product_key = 'SymbolList['
    SymbolList = {}
    for j in range(len(work_articles_list)):
        if work_articles_list[j] == None:
            print('NONE!!!')
            break
        SymbolList[product_key+str(j)+']'] = work_articles_list[j]
    print(SymbolList)
    
    CountryLanguageParams={'Country' : 'RU','Language' : 'RU',}
    params={**CountryLanguageParams,**SymbolList}
    print(params)
    token = 'TOKEN'
    app_secret = 'APP SECRET'
    action = 'Products/GetProducts' # request method
    all_data = product_import_tme(token, app_secret, action, params) 
    for j in range(len(work_articles_list)):
        print('Артикул: '+SymbolList['SymbolList['+str(j)+']'],'\nкраткое описание: ', all_data['Data']['ProductList'][j]['Description'],'\nВес: ', all_data['Data']['ProductList'][j]['Weight']*0.001)
    time.sleep(0.1)
    return get_data((new_cord[1]+1))
 
get_data(6)

wb.Close()
Excel.Quit()