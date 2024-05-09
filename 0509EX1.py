# -*- coding: utf-8 -*-
"""
Created on Thu May  9 19:31:50 2024

@author: USER
"""

# break 立即跳脫
# continue 跳下一個敘述

# for i in range(100):
#     if i == 15:
#         break
#     if i % 3 ==0:
#         continue
#     print("i=",i)
#     print("平方:",i*i)
# print("FINISH")

# for i in range(10):
#     for a in range(100):
#         if a ==5:
#             break
#         print("i=",i,"a=",a)
#     print("F")

# ans = 72
# guess=0
# count =1
# while ans != guess:
#     guess = int(input("輸入:1~100之間整數:"))
#     if guess >ans:
#        print("請猜小一點")
#     elif guess < ans :
#        print("請猜大一點")
#     count += 1
#     if  count == 4:
#          break
# if ans == guess:
#        print("RIGHT")   
# else:
#         print("次數已滿")

import random
ans = random.randint(1, 100)
guess=0
count =1
while ans != guess:
    guess = int(input("輸入:1~100之間整數:"))
    if guess >ans:
       print("請猜小一點")
    elif guess < ans :
       print("請猜大一點")
    count += 1
    if  count == 4:
         break
if ans == guess:
       print("RIGHT")   
else:
        print("次數已滿")