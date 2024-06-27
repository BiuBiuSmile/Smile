# -*- coding: utf-8 -*-
"""
Created on Thu Jun 27 19:10:42 2024

@author: USER
"""

import matplotlib.pyplot as plt
data1 = [1,2,3,4,5,6,7,8]
data2 = [1,4,9,16,25,36,49,64]
data3 = [1,3,6,10,15,21,28,37]
data4 = [1,7,15,26,40,57,77,100]
seq=[1,2,3,4,5,6,7,8]

plt.plot(seq,data1,'g-*',seq,data2,'r-o',seq,data3,'y-^',seq,data4,'k-s')
plt.title('Chart')
plt.xlabel('X_Value')
plt.ylabel('Y_Value')
plt.tick_params(axis='both',labelsize=14,color='red')
plt.show()