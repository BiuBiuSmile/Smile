# -*- coding: utf-8 -*-
"""
Created on Tue Jun 11 19:49:43 2024

@author: USER
"""

class Father:
    def play(self,item):
        print("使用",item,"來玩")
class Son:
    def go(self):
        print("bike")
son=Son()
son.go()
Father.play(son, '球棒')

class Mother():
    def display(self,pay):
        self.price=pay
        if self.price >=30000:
            self.price *=0.9
        print('={:,}'.format(self.price))
class Daughter(Mother):
    def display(self,pay):
        self.price=pay
        super().display(pay)
        if self.price >=30000:
            self.price *= 0.8
        print('8折{:,}'.format(self.price))
Mary=Mother()
print('40000打9折=',end="")
Mary.display(40000)


cherry=Daughter()
print('30000打9折=',end="")
cherry.display(30000)
    
        
        
        
        
        