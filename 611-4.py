class Parent:
    def show1(self):
        print("Parent show1")
    def show2(self):
        print("Parent show2")
class Son(Parent):
    def display(self):
        print('Son display')
class Daughter(Parent):
    def show2(self):
        print("Daughter show2")
    def display(self):
        print("Daughter display")
class GrandChild(Son,Daughter):
    def message(self):
        print('Child msg')
bill=GrandChild()
bill.show1()
bill.show2()
bill.display()
bill.message()










