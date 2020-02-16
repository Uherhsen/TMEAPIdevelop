# -*- coding: utf-8 -*-
"""
Created on Sun Feb  2 22:26:59 2020
регулярка для отделения по точке-запятой первых слов из декрипшина
Далее это будет называться ТИПОМ продукта
"""
import re

def product_name(descr):
    product = re.split(r';', descr)
    return(product[0])

if __name__ == '__main__':
    test_descr = 'Вентилятор: DC; осевой; 12ВDC; 38x38x20мм; 20,39м3/ч; 44дБА; 26AWG'
    print(product_name(test_descr))