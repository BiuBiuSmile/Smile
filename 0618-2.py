# -*- coding: utf-8 -*-
"""
Created on Tue Jun 18 19:21:08 2024

@author: USER
"""

import datetime
today = datetime.datetime.today()
print(today)
today2=datetime.date.today()
print(today2)
f=datetime.datetime.strftime(today,'%Y%m%d%H%M%S')
print(f)
f2=datetime.datetime.strftime(today,'%y%m%d')
print(f2)