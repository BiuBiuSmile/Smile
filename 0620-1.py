# -*- coding: utf-8 -*-
"""
Created on Thu Jun 20 19:17:31 2024

@author: USER
"""

import csv
fn="member.csv"
with open(fn,encoding=('utf-8')) as fObj:
    csvReader = csv.reader(fObj)
    for row in csvReader:
      print("Row %s = " % csvReader.line_num,row)
      
      for item in row:
          print(item)
          

    