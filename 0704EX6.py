#用迴圈將多餘的數字刪掉


data=[1,2,3,'a','b','b',1]
temp = list()
for i in data:
    if not(i in temp):
        temp.append(i)
print(temp)
