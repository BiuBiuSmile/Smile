# -*- coding: utf-8 -*-
"""
Created on Tue May 28 19:21:10 2024

@author: USER
"""
def total(*value):
    t=0
    for i in value:
        t += i
    return t
print(total())
print(total(1,2,3,4))
print(total(10,20))

def school(*people,name):
    print("校名:",name)
    print("學校資料:",people)
#school(300, 100,2,"很友愛學校")錯誤的
school(300,100,2, name="很友愛學校")