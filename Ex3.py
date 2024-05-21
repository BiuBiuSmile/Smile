# -*- coding: utf-8 -*-
"""
Created on Tue May 21 20:04:50 2024

@author: USER
"""

fruits={'apple':100,'banana':49,'orange':69}
print(fruits['orange'])
# print(fruits['cherry']) 因字典沒有這個詞所以會發生錯誤
print(fruits.get('cherry',0)) #尋找是否有這個詞，沒有的話會顯示None
print(fruits.get('cherry','找不到'))
print(fruits.get('orange','找不到'))