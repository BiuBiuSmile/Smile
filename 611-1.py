class Motor:
    def __init__(self,name,price=95,capacity=1800):
        self.name=name
        self.price=price
        self.capacity=capacity
    def equip(self,award):
        self.price+=award
    def __str__(self):
        msg="{0:8s},售價:{1:7.2f}萬元,排氣量:{2:,}CC數"
        return msg.format(self.name,self.price,self.capacity)
    
    
class Hybrid(Motor):
    def equip(self, award,cell=3.64):
        Motor.equip(self, award+cell)
    def tinted(self,color):
        if color == 'r':
            return  "大方紅"
        else :
            return "神秘黑"
stand = Motor("stand")
apollo=Motor("Apollo",price=81.2,capacity=1500)
print(apollo,"不含電子防盜鎖") 
apollo.equip(1.9)       

inno=Hybrid("Innovate",360.12)
inno.equip(1.3)
print("Hybrid color is",inno.tinted('r'))
print("三種車款")

for item in (stand,apollo,inno):
    print(item)







        
        
        
        