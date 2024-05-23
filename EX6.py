# -*- coding: utf-8 -*-
"""
Created on Thu May 23 21:11:14 2024

@author: USER
"""

def area(w,h):
    return w*h
def loopfor(start,end):
    total=0
    for i in range(start,end+1):
        total +=i
        return total
print(area(3, 4))
print(loopfor(2,30))