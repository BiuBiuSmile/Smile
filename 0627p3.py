# -*- coding: utf-8 -*-
"""
Created on Thu Jun 27 19:18:01 2024

@author: USER
"""

import matplotlib.pyplot as plt
BMW=[4100,3599,5700]
Benz=[3399,5100,4166]
MG=[2100,7000,8100]
seq=[2021,2022,2023]
plt.xticks(seq)
lineBMW,=plt.plot(seq,BMW,'-o',label='BMW')
lineBenz,=plt.plot(seq,Benz,'-*',label='Benz')
lineMG,=plt.plot(seq,MG,'-^',label='MG')

plt.legend(handles=[lineBMW,lineBenz,lineMG],loc='best')#loc預設為best，1是右上，2是左上，3是左下
plt.title('car sales')
plt.xlabel('Year')
plt.ylabel('Number')
plt.show()
