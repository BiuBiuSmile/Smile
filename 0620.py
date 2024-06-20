# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import csv
fn="member.csv"
with open(fn,encoding=('utf-8')) as fObj:
    csvReader = csv.reader(fObj)
    data=list(csvReader)
print(data)