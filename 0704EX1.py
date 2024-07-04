import sqlite3
conn = sqlite3.connect('web.db')

sql="create table students(sid integer primary key autoincrement,name carchar(30),sex varchar(2),address varchar(100))"#autoincrement是自動增值
cusrsor = conn.cursor()
cusrsor.execute(sql)
conn.commit() #會立即更新資料表
conn.close() #關閉資料表