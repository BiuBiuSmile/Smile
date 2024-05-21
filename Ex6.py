# -*- coding: utf-8 -*-
"""
Created on Tue May 21 21:17:59 2024

@author: USER
"""

student={'周子瑜':92,'IU':89}
name = input('輸入學生姓名:')
if name in student:
    print(name,'的成績:',student[name])
else:
    score = int(input('輸入分數:'))
    student[name]=score
print(student)
keys=student.keys()
value=student.values()
print(keys)
print(value)

items=student.items()
print(list(items))

it = list(items)
print(it[0])
print(it[0][0])
print(it[0][1])

for k,v in studens.items():
    print(k)
    print(v)
    print()






