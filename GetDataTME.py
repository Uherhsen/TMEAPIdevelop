# -*- coding: utf-8 -*-
"""
Created on Tue Jan 14 12:02:49 2020

Получение данных о электронных компонентах с сайта TME через его API
"""
import win32com.client,time
from TME_Python_API import product_import_tme

Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(u"productdata.xlsx") #путь к файлу
sheet = wb.ActiveSheet

i = 1
sumcells = ""
while sheet.Cells(i,1).value != None:
    i+=1
       
cord = 'A'+str(1)+':A'+str((i-1)) 

work_articles_list = [r[0].value for r in sheet.Range(cord)]

params={'Country' : 'RU','Language' : 'RU',}
token = 'ac434c181917ed4e51c49a2027bfd040e9f2da0054be7'
app_secret = '0b748f6e5d340d693703'
action = 'Products/Search' # request method, метод пинг Utils/Ping

def ping():
        ping_data = product_import_tme(token, app_secret, 'Utils/Ping', params={})
        print(ping_data)
               
def search_articles(rng1=0,rng2=len(work_articles_list)):
    for j in range(rng1,rng2):
        params['SearchPlain'] = str(work_articles_list[j])
        all_data = product_import_tme(token, app_secret, action, params)
        try :
            print(all_data['Data']['ProductList'][0]['Symbol'],all_data['Data']['ProductList'][0]['Description'])
            sheet.Cells(j+1,2).value = all_data['Data']['ProductList'][0]['Description']
            weight = (all_data['Data']['ProductList'][0]['Weight'])*0.001
            sheet.Cells(j+1,3).value = weight
        except IndexError:
            if all_data["Status"]=="OK":
                print('\nСтатус сети ',all_data["Status"],'\nАртикул "'+work_articles_list[j]+'" отсутствует на TME\n')
                sheet.Cells(j+1,2).value = 'Артикула нет на TME'
            else:
                print('\nСтатус сети ',all_data["Status"])
                sheet.Cells(j+1,2).value = all_data['Data']['ProductList'][0]['Description']
                time.sleep(2) 
            continue
 
search_articles()

wb.Close()
#Excel.Quit()

# Errors:
# HTTPError: Bad Request
#IndexError: list index out of range

