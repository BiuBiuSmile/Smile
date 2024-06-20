# -*- coding: utf-8 -*-
"""
Created on Thu Jun 20 20:08:04 2024

@author: USER
"""

import csv
fn="text.csv"

data=[['A'],['B'],['C']]
     
with open(fn,'w',encoding="utf-8",newline='') as fObj:
    
    csvWriter=csv.writer(fObj)
    csvWriter.writerows(data)
    