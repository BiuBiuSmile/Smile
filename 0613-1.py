file="yahoo.txt"
file_Obj=open(file,encoding="utf8")
data=file_Obj.read()
file_Obj.close()
print(data)