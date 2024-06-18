# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

file = "lcc.txt"
word="Hello Python第一階"
i=100
with open(file,'a',encoding="utf-8") as fObj:#a是續寫的意思，如果改成w會變成覆寫
      fObj.write(word+"\n")
      fObj.write(str(i))
      