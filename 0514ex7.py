# -*- coding: utf-8 -*-
"""
Created on Tue May 14 19:57:34 2024

@author: USER
"""
#count=計算次數
words=['a','b','c','d','a','c','f']
a=words.count('a')
f=words.count('f')
g=words.count('g')
print('a有:',a)
print('f有:',f)
print('g有:',g)
#找尋索引位置
ind=words.index('d')
print("d的索引位置:",ind)


start=0
for i in range(words.count('c')):
    ind=words.index('c',start)
    print('c的索引:',ind)
    start=ind+1