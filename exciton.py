#import numpy as np
#import math
#from openpyxl import Workbook
#from openpyxl import load_workbook
#from openpyxl.chart import (ScatterChart, Reference, Series)
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import time

from ClassExciton import Exciton


#ex = Exciton('HamiltonBrixner-6lower.xlsx','HamiltonBrixner-6lower-out.xlsx')
ex = Exciton('HamiltonBrixner2.xlsx','HamiltonBrixner2.xlsx')

#ex.calculate()
t1 = time.time()
for i in range(0,100):
    ex.getsticks()
    ex.getspectra()
t2 = time.time()-t1
print("Time = ",t2/100)
#ex.wb.save('test.xlsx')
fig, (ax1,ax2) = plt.subplots(2,1, sharex=True)
ax1.plot(ex.x, ex.abs)
ax2.plot(ex.x,ex.CD)

ax1.set(xlabel='', ylabel='absorbance',title='Exciton spectrum')
ax2.set(xlabel='wavenumers', ylabel='CD')
ax1.grid()
ax2.grid()

#fig.savefig("test.png")
plt.show()
