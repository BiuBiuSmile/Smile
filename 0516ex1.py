# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

score= [99,89,92,100,62,73]
data=sorted(score)
print('舊資料',score)
print("排序後的資料:",data)
print(sorted(score,reverse=True))
for i in sorted(score,reverse=True):
    print(i)
    