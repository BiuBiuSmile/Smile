# -*- coding: utf-8 -*-
"""
Created on Thu Jun 13 21:33:01 2024

@author: USER
"""

file="yahoo.txt"
with open(file,encoding="utf8") as fo:
     data=fo.readlines()

i=1
for row in data:
    if "Doncic"in row:
        print("第{}行找到".format(i))
        print(row)
    i+=1