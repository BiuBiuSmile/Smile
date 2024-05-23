# -*- coding: utf-8 -*-
"""
Created on Thu May 23 19:11:54 2024

@author: USER
"""

data=list()
for i in range(1,11):
    data.append(i)
print(data)


level=[i for i in range(1,11)]#串列表達式
print(level)

number=[i for i in range(1,11) if i % 3 ==0]
print(number)