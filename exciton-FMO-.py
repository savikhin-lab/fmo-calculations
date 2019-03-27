#import numpy as np
#import math
#from openpyxl import Workbook
#from openpyxl import load_workbook
#from openpyxl.chart import (ScatterChart, Reference, Series)
import matplotlib.pyplot as plt
import numpy as np

from ClassExciton import Exciton


nam = "HamiltonBrixner"; N=8
#nam = "HamiltonKell"; N=9

ex = np.ndarray(N,Exciton)
namin = nam+".xlsx"
for i in range(0,N):    # knockout pigment i, i=0 - full hamiltonian
    namout = nam+str(i)+".xlsx"
    ex[i] = Exciton(namin,namout,i) # knockout pigment i
    ex[i].bandwidth = 200           # overwrite spectral bandwidth before calculating
    ex[i].calculate()
    if i==0:
        x = np.zeros(len(ex[0].x))
        x = 1.e7/ex[0].x            # convert wavenumbers into nanometers
    namabs = nam+"ABS"+str(i)+".xy"# names of xy files for spectra, easier to load into SpectraSolve
    namcd = nam+"CD"+str(i)+".xy"
    file_abs = open(namabs,'w')
    file_cd = open(namcd,'w')
    for j in range(0,len(x)):
        file_abs.write("%g,%g\r\n" % (x[j],ex[i].abs[j]))
        file_cd.write("%g,%g\r\n" % (x[j],ex[i].CD[j]))
    file_abs.close()
    file_cd.close()

absd = np.zeros((N,len(x)))
cdd = np.zeros((N,len(x)))
absd[0] = ex[0].abs
cdd[0] = ex[0].CD
# make difference spectra for knockout pigments
for i in range(1,N):
    absd[i] = ex[i].abs - ex[0].abs
    cdd[i] =  ex[i].CD - ex[0].CD

# next will plot whatever we set a and c to be (combinations of knockouts, for example)
a = absd[2] + 0.*absd[2]
c = cdd[2] - 0.*cdd[2]
#ex.wb.save('test.xlsx')
fig, (ax1,ax2) = plt.subplots(2,1, sharex=True)
ax1.plot(x,a)
ax2.plot(x,c)

ax1.set(xlabel='', ylabel='absorbance',title='Exciton spectrum')
ax2.set(xlabel='wavenumers', ylabel='CD')
ax1.grid()
ax2.grid()

#fig.savefig("test.png")
plt.show()
