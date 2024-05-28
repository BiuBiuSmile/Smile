def circle(r=5):
    print('半徑:',r)
    area = r*r*3.14
    print('圓面積:',area)
circle()
circle(11)

def city(number,name='台中',parent='市民'):
    #函式裡面有預設值後面都要是預設值
    print(number)
    print(name)
    print(parent)
city(400)
city(700,'高雄')
city(900,'台東','原住民')