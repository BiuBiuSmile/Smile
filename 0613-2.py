# -*- coding: utf-8 -*-
"""
Created on Thu Jun 13 21:19:18 2024

@author: USER
"""

file="yahoo.txt"
with open(file,encoding="utf8") as fo:
    data=fo.read()
    print(data.rstrip())