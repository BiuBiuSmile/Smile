# -*- coding: utf-8 -*-
"""
Created on Tue Jun 11 21:26:58 2024

@author: USER
"""

class Father:
    def display(self,name):
        self.name=name
        print("Father name is",self.name)
class Mother:
    def display(self,name):
        self.name=name
        print("Mother name is",self.name)
class Child(Father,Mother):
    pass
class Son(Father):
    pass
print(Child.__name__,"類別,繼承兩個類別")

for item in Child.__bases__:#動態類別
     print(item)
john=Son()
john.display("Tom")
print("son的父類:",Son.__bases__)
Son.__bases__=(Mother,)
john.display("Mary")


