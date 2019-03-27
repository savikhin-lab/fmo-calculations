import numpy as np
import math
#from openpyxl import Workbook
#from openpyxl import load_workbook
#from openpyxl.chart import (ScatterChart, Reference, Series)
import random as random
import matplotlib.pyplot as plt


from ClassExciton1 import Exciton



noise = 100 # diagonal disorder bandwidth in energy

ex = Exciton('HamiltonBrixner.xlsx')
ex.bandwidth = 20
ex.recalculate()


#ex.save_excel('DimerOut.xlsx')

totabs = np.zeros((len(ex.x)))
totCD = np.zeros((len(ex.x)))

totabs = ex.abs
totCD = ex.CD

N=ex.size

#vary diagonal energies:
#x=random.choices()
if noise>0:
    sig = noise/math.sqrt(8.*math.log(2.))
    for i in range(0,1000):
        for j in range(0,N):
            r=random.gauss(0,sig)
            ex.ham[j,j] = ex.origham[j,j]+r
        ex.recalculate()

ex.save_excel('TrimerOut.xlsx')
#r = random.gauss(0,sig) #normal noise


fig, (ax1,ax2) = plt.subplots(2,1, sharex=True)
ax1.plot(ex.x,ex.abs)
ax2.plot(ex.x,ex.CD)

ax1.set(xlabel='', ylabel='absorbance',title='Exciton spectrum')
ax2.set(xlabel='wavenumers', ylabel='CD')
ax1.grid()
ax2.grid()

#fig.savefig("test.png")
plt.show()
