class Father:
    def __init__(self,name):
        self.name=name
        self.money=0
    def display(self):
        print(self.name,self.money)
F=Father("Bill")

F.car="BMW"
print(F.car)
print(F.name)
a=Father("Tom")
print(a.name)
