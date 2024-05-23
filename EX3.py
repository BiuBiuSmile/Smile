# -*- coding: utf-8 -*-
"""
Created on Thu May 23 19:24:47 2024

@author: USER
"""

students={'John':40,'Peter':70,'Mary':51,'Eric':43,'Bill':88}
up60={K:V for K,V in students.items() if V >=60}
low60={K:V for K,V in students.items() if V <60}
print('不及格有:',len(low60))
for K,V in low60.items():
    print("%-10s%3d"%(K,V))