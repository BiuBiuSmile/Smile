# -*- coding: utf-8 -*-
"""
Created on Tue Jun 18 21:18:07 2024

@author: USER
"""

import os
path=os.path.join("c:\\","lcc")
for file in os.listdir("c:\\lcc"): #listdir列出指定目錄中的所有文件和子目錄的名稱
    print(file)
    f=os.path.join(path,file)
    if os.path.isdir(f): #isdir用於檢查指定路徑是否為一個目錄（資料夾）
        print("是目錄")
    elif os.path.isfile(f): #isfile用於檢查指定路徑是否為一個檔案
        print("是檔案")