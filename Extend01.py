# -*- coding: utf-8 -*-
"""
Created on Thu Jun  6 21:03:47 2024

@author: USER
"""

class Father:
    def company(self):
        print("老爸公司:輝達")
    def showMoney(self):
        print("有3兆美元")
class  Son(Father):
    pass
class Daughter(Father):
    def company(self):
        print("在超微工作")
    def boyfriend(self):
        print("在intel服務")
        
son=Son()
daughter=Daughter()
son.company()
son.showMoney()
daughter.company()
daughter.showMoney()