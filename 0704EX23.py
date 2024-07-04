import sqlite3
conn = sqlite3.connect('web.db')
cursor=conn.cursor()
#sql="select name,sex from students"
sql="select * from students"
#sql="select name,sex from students where sex='M'"
cursor.execute(sql)

data = cursor.fetchall()
for row in data:
    for item in row:
        print(item,end=',')
    print()
print(data)
conn.close()
