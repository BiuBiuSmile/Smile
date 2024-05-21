# -*- coding: utf-8 -*-
"""
Created on Tue May 21 20:27:55 2024

@author: USER
"""

student={'John':81,'Peter':62}
student['Bill']=100
print(student)

student.update({'Mary':83,'Sue':92})
print(student)

print('排序')
for k in sorted(student):
    print('%-12s %3d'%(k,student[k]))
student.pop('John')

for k in sorted(student,reverse=True):
    print('%-12s %3d'%(k,student[k]))
    
print("清空字典",student.clear())
student.update(Eric=92,George=73)
print(student)
