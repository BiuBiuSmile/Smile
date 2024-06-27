# -*- coding: utf-8 -*-
"""
Created on Thu Jun 27 20:10:34 2024

@author: USER
"""

import matplotlib.pyplot as plt

import numpy as np
x=np.linspace(0,5,500)
y=1-0.5*np.abs(x-2)
lwidth=(1+x)**2
plt.scatter(x,y,s=lwidth,color='g')
plt.show()