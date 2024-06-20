# -*- coding: utf-8 -*-
"""
Created on Thu Jun 20 19:30:07 2024

@author: USER
"""

import csv
fn="member.csv"
with open(fn,encoding="utf-8") as fObj:
    csvDic=csv.DictReader(fObj)
    for row in csvDic:
        print(row)
        print(row['sex'])
        print(row['name'])