# -*- coding: utf-8 -*-
"""
Created on Tue May 28 19:25:08 2024

@author: USER
"""

def student(name,*score,subject=5):
    if subject >=1:
        print("姓名:",name)
        print("共有",subject,"科分數:",score)
    total=sum(score)
    print("總分:",total,"平均:%.3f"%(total/subject))
student("Mary", 63,73,82,100,70)
student("john", 81,92,33,subject=3)
