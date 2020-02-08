# -*- coding: utf-8 -*-
"""
Created on Tue Jan 14 12:02:49 2020

Получение данных о электронных компонентах с сайта TME через его API
"""
import win32com.client,time
from TME_Python_API import product_import_tme

# Функция считает все ячейки в которых что то написано,до тех пор пока не встретит пустую ячейку "None"
def number_of_articles():
    # Открываем Эксель
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(u"d:\\usr\\documents\\Desktop\\TMEAPIdevelop\\APP\\productdata.xlsx") #путь к файлу   
    sheet = wb.ActiveSheet
    #выясняем количество артикулов в файле эксель
    i = 1
    while sheet.Cells(i,1).value != None:
        i+=1
    wb.Close()
    Excel.Quit()
    time.sleep(1)
    return i-1

# Функция создающая список артикулов. Принимает число артикулов и номер колонки в виде буквы (str): 'A'- первая колонка
    
def articles_list(n,column):
    # Открываем Эксель
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(u"d:\\usr\\documents\\Desktop\\TMEAPIdevelop\\APP\\productdata.xlsx") #путь к файлу   
    sheet = wb.ActiveSheet
    cord = column+str(1)+':'+column+str((n)) 
    # формирование списка артикулов
    return [r[0].value for r in sheet.Range(cord)]
    wb.Close()
    Excel.Quit()
    time.sleep(1)
    
# Функция для вывода пинга
    
def ping():
        ping_data = product_import_tme(token, app_secret, 'Utils/Ping', params={})
        print(ping_data)
        
# Функция проставляет оригинальные артикулы производлтеля, дескрипшен, ссылку на фото, ссылку на страницу продукта и вес в КГ! 
            
def search_articles(articles_list, rng1=0):
    '''Поиск базовых параметров'''
    # Открываем Эксель
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(u"d:\\usr\\documents\\Desktop\\TMEAPIdevelop\\APP\\productdata.xlsx") #путь к файлу   
    sheet = wb.ActiveSheet
    
    rng2=len(articles_list)
    for j in range(rng1,rng2):
        params['SearchPlain'] = str(articles_list[j])
        all_data = product_import_tme(token, app_secret, action1, params)
        #print(all_data)
        try :
            print(all_data['Data']['ProductList'][0]['Symbol'],all_data['Data']['ProductList'][0]['Description'])
            all_data['Data']['ProductList'][0]['Symbol']
            sheet.Cells(j+1,2).value = all_data['Data']['ProductList'][0]['Symbol']
            sheet.Cells(j+1,3).value = all_data['Data']['ProductList'][0]['Description']
            sheet.Cells(j+1,5).value = all_data['Data']['ProductList'][0]['Photo'][2:]
            sheet.Cells(j+1,6).value = all_data['Data']['ProductList'][0]['ProductInformationPage'][2:]
            weight = (all_data['Data']['ProductList'][0]['Weight'])
            if all_data['Data']['ProductList'][0]['WeightUnit']=='g':
                weight = weight*0.001
                #print('{:f}'.format(weight))
                weight = round(weight,(('{:f}'.format(weight)).count('0'))+1) # округление
            sheet.Cells(j+1,4).value = weight
        except IndexError:
            if all_data["Status"]=="OK":
                print('\nСтатус сети ',all_data["Status"],'\nАртикул "'+articles_list[j]+'" отсутствует на TME\n')
                sheet.Cells(j+1,3).value = 'Артикула нет на TME'
            else:
                print('\nСтатус сети ',all_data["Status"])
                sheet.Cells(j+1,3).value = all_data['Data']['ProductList'][0]['Description']
                time.sleep(2) 
            continue
    wb.Close()
    Excel.Quit()
    time.sleep(1)
    print('\nГотово')
    
# функция использует "экшен" для поиска параметров, к которому нужны оригинальные артикулы,
# запускается только после функции search_articles
    
def search_param(articles_list,rng1=0,):
    # Открываем Эксель
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(u"d:\\usr\\documents\\Desktop\\TMEAPIdevelop\\APP\\productdata.xlsx") #путь к файлу   
    sheet = wb.ActiveSheet
    rng2=len(articles_list)
    print('\nЦикл проставления параметров\n')
    for j in range(rng1,rng2):
        if articles_list[j] != None:
            params['SymbolList[0]'] = articles_list[j]
            all_data = product_import_tme(token, app_secret, action2, params)
            if all_data['Status'] == "OK":
                print(all_data['Data']['ProductList'][0]['ParameterList'][1]['ParameterName'],
                      all_data['Data']['ProductList'][0]['ParameterList'][1]['ParameterValue'])
                prms={}
                for i in all_data['Data']['ProductList'][0]['ParameterList']:
                    prms[i['ParameterName']] = i['ParameterValue']
                sheet.Cells(j+1,7).value = str(prms)
                
                time.sleep(0.1)
            else:
                print('ошибка статуса сети')
                j+=1
        else:
            print('Пропуск артикула')
            j+=1
    wb.Close()
    Excel.Quit()
    time.sleep(1)
    print('\nГотово')

def products_files(articles_list,rng1=0):
    # Открываем Эксель
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(u"d:\\usr\\documents\\Desktop\\TMEAPIdevelop\\APP\\productdata.xlsx") #путь к файлу   
    sheet = wb.ActiveSheet
    rng2=len(articles_list)
    print('\nЦикл проставления ссылок на даташит\n')
    for j in range(rng1,rng2):
        if articles_list[j] != None:
            params['SymbolList[0]'] = articles_list[j]
            all_data = product_import_tme(token, app_secret, action3, params)
            if all_data['Status'] == "OK":
                try:
                    print(all_data['Data']['ProductList'][0]['Files']['DocumentList'][0]['DocumentUrl'][2:],'\n'+('_'*50)) #[0]['DocumentUrl'])
                    sheet.Cells(j+1,8).value =str(all_data['Data']['ProductList'][0]['Files']['DocumentList'][0]['DocumentUrl'][2:])
                except KeyError:
                    print('KeyError')
                    j+=1
                except IndexError:
                    print('Нет PDF:\n',all_data['Data']['ProductList'][0]['Files'],'\n'+('_'*50))
                    j+=1
            else:
                print('ошибка статуса сети')
                break
                
        else:
            print('Нет артикула','\n'+('_'*50))
            j+=1
    print('\nГотово')
    time.sleep(1)
    wb.Close()
    Excel.Quit()
          
    
params={'Country' : 'RU','Language' : 'RU',}
token = 'TOKEN'
app_secret = 'APP SECRET'
action1 = 'Products/Search' # request method, метод пинг Utils/Ping
action2 = 'Products/GetParameters'
action3 = 'Products/GetProductsFiles'
 

n = number_of_articles()

work_articles_list_A = articles_list(n,column='A')    
search_articles(work_articles_list_A)
work_articles_list_B = articles_list(n,column='B')
search_param(work_articles_list_B)
products_files(work_articles_list_B)