# -*- coding: utf-8 -*-
"""
Created on Tue Jun 18 21:02:44 2024

@author: USER
"""

import os 
path = os.path.join("C:\\", "demo","file")
if not os.path.exists(path):
    os.makedirs(path,exist_ok=True)
    print("建立成功")