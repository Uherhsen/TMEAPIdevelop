# -*- coding: utf-8 -*-
"""
Created on Tue Jan 14 12:02:49 2020

Получение данных о электронных компонентах с сайта TME через его API

Извлеч ПДФ! РЕФАКТОРИНГ
"""
import win32com.client,time
from TME_Python_API import product_import_tme
             
def search_param(articles_list,rng1=0,):
    rng2=len(articles_list)
    for j in range(rng1,rng2):
        if articles_list[j] != None:
            params['SymbolList[0]'] = articles_list[j]
            all_data = product_import_tme(token, app_secret, action, params)
            if all_data['Status'] == "OK":
                print(all_data['Data']['ProductList'][0]['ParameterList'])
                sheet.Cells(j+1,7).value =str(all_data['Data']['ProductList'][0]['ParameterList'])
                time.sleep(0.2)
            else:
                j+=1
        else:
            j+=1
    print('\nГотово')    

params={'Country' : 'RU','Language' : 'RU',}
token = 'ac434c181917ed4e51c49a2027bfd040e9f2da0054be7'
app_secret = '0b748f6e5d340d693703'
action = 'Products/GetParameters' # request method, метод пинг Utils/Ping    

    
# Открываем Эксель
Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(u"d:\\usr\\documents\\Desktop\\TMEAPIdevelop\\APP\\productdata.xlsx") #путь к файлу   
sheet = wb.ActiveSheet
#выясняем количество артикулов в файле эксель
i = 1
while sheet.Cells(i,1).value != None:
    i+=1       
cord = 'B'+str(1)+':B'+str((i-1)) 
# формирование списка артикулов
work_articles_list2 = [r[0].value for r in sheet.Range(cord)]

search_param(work_articles_list2)
wb.Close()
Excel.Quit()

#if __name__ == "__main__":
    