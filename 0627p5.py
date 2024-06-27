import matplotlib.pyplot as plt
x=list(range(1,101))
y=[i ** 2 for i in x]
plt.scatter(x,y,color='y')
plt.show()