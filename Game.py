# -*- coding: utf-8 -*-
"""
Created on Thu Jun 13 18:54:51 2024

@author: USER
"""

from Role import Avisder,SworksMan,Dancer
import random
def fight(me,you):
    print(me.getName(),end='')
    if isinstance(me,Avisder):
        if me.getHp()<=50:
            me.cure()
            me.setHp(100)
            print(me.getName(),'的血量',me.getHp())
        else:
            me.fight()
            n=random.randint(1,10)
    
            if n ==3 or n==5 or n==7:
                print(you.getName(),'miss')
            else:
                blood = random.randint(0,10)
                youblood = you.getHp() - blood
                you.setHp(youblood)
                print("{}損失{}血量，目前的血量：{}".format(you.getName(),blood,youblood))
    else:
        me.fight()
        n=random.randint(1,10)
    
        if n ==3 or n==5 or n==7:
                print(you.getName(),'miss')
        else:
            blood = random.randint(0,10)
            youblood = you.getHp() - blood
            you.setHp(youblood)
            print("{}損失{}血量，目前的血量：{}".format(you.getName(),blood,youblood))


if __name__=='__main__':
    com=list()
    player=list()
    com.append(Avisder("司馬懿",100,90))
    com.append(SworksMan("曹操",100,81))
    com.append(Dancer("甄姬",100,60))
    player.append(Dancer("小喬",100,90))
    player.append(SworksMan("孫權",100,60))
    player.append(Avisder("周瑜",100,90))
    
while (len(com)>0 and len(player)>0):
    f=random.randint(1,100)
    if f %2 ==0:
        n=random.randint(0,len(com)-1)
        minblood=list()
        for r in player:
            minblood.append(r.getHp())
        selplayer=player[(minblood.index(min(minblood)))]       
        fight(com[n],selplayer)
        
        if selplayer.getHp()<=0:
            player.remove(selplayer)
    else:
        n=random.randint(0,len(player)-1)
        minblood=list()
        for r in com:
            minblood.append(r.getHp())
        selplayer=com[(minblood.index(min(minblood)))]  
        fight(player[n],selplayer)
        if selplayer.getHp()<=0:
            com.remove(selplayer)
if len(com)>0:
    print("com win")
else:
    print("玩家贏")
       
        
        
        