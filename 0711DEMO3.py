# -*- coding: utf-8 -*-
"""
Created on Thu Jul 11 21:29:04 2024

@author: USER
"""

import sqlite3
conn=sqlite3.connect("web.db")
cursor=conn.cursor()
#sql="select students.name,lesson.lessonname,lesson.score from students inner join lesson on students.sid = lesson.sid"
sql="select students.name,lesson.lessonname,lesson.score from students left join lesson on students.sid = lesson.sid where lesson.lessonname is null"
cursor.execute(sql)
data=cursor.fetchall()
for row in data:
    for item in row:
       print(item,end="\t")
    print()