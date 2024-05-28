# -*- coding: utf-8 -*-
"""
Created on Tue May 28 21:24:57 2024

@author: USER
"""

score=[61,31,50,70,90,72]
print('最小值:',min(score))
print('最大值:',max(score))
print('加總:',sum(score))

print(divmod(11, 2))
print(abs(-100))
print(float("1.234"))
print(pow(2, 3))
print(round(100.345,1))

name=['Bill','Mary','Peter']
score=[92,100,69]
item=zip(name,score)
print(list(item))
for k,v in zip(name,score):
    print(k)
    print(v)