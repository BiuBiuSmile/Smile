from Role import Dancer,SworksMan
import random

def changeRole(me,you):
     print(me.getName(),end='')
     me.fight()
     blood=random.randint(0, 10)
     youblood=you.getHp()-blood
     you.setHp(youblood)
     print("{}損失{}血量,目前血量:{}".format(you.getName(),blood,youblood))



if __name__ == '__main__':
    dancer=Dancer('大喬',150,300)
    man = SworksMan('劉備',150,100)
    
    while(dancer.getHp()>0 and man.getHp()>0):
        num=random.randint(1,50)
        if num%2==0:
           changeRole(dancer,man)
        else:
           changeRole(dancer,man)
    if dancer.getHp()<=0:
            print(man.getName(),"win")
    else:
            print(dancer.getName(),"Win")