import sqlite3

conn = sqlite3.connect('web.db')

sql = """
create table lesson(
    id integer primary key autoincrement,
    sid int,
    lessonname varchar(50),
    score int)
"""

cursor = conn.cursor()  # 建立一個資料庫連線資料集

cursor.execute(sql)

conn.commit()  # 立即提交 ，將緩沖區的資料即時寫入到資料庫


conn.close()