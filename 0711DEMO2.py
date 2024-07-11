# -*- coding: utf-8 -*-
"""
Created on Thu Jul 11 21:14:59 2024

@author: USER
"""

import sqlite3

conn = sqlite3.connect('web.db')

cursor = conn.cursor()

sql = """
insert into lesson(sid,lessonname,score) values(2,'python',3),(4,'AI',1),(5,'python',3)

"""

cursor.execute(sql)

conn.commit()

conn.close()