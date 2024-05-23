# -*- coding: utf-8 -*-
"""
Created on Thu May 23 21:24:59 2024

@author: USER
"""

def circle(r=10):
    area = r*r*3.14
    circleLen=2*r*3.14
    return area,circleLen
a,c=circle()
ac=circle(5)
print(ac)