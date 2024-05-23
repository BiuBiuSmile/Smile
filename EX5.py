# -*- coding: utf-8 -*-
"""
Created on Thu May 23 20:51:39 2024

@author: USER
"""
#無參數的函式呼叫
def Hello():
    print("你好")
    print("Hello")
def loopfor():
    for i in range(5):
        print(i)
Hello()
Hello()
loopfor()

def sumNumber():
    total=0
    for i in range(1,11):
        total +=i
    return total
number=sumNumber()
print(number)
print(sumNumber())
