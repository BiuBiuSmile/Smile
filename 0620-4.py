# -*- coding: utf-8 -*-
"""
Created on Thu Jun 20 19:59:35 2024

@author: USER
"""

import csv
fn="bike.csv"
with open(fn,'w',encoding="utf-8",newline='') as fObj:
    
    csvWriter=csv.writer(fObj)
    csvWriter.writerow(['sna','sbi','bemp'])
    csvWriter.writerow(['三重路口','10','31'])
    csvWriter.writerow(['總統府','30','2'])