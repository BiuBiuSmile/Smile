# -*- coding: utf-8 -*-
"""
Created on Thu Jul  4 21:19:33 2024

@author: USER
"""
#刪除
import sqlite3
conn = sqlite3.connect('web.db')
cursor=conn.cursor()
sql="delete from students where sid = 1"

cursor.execute(sql)
conn.commit()
conn.close()
