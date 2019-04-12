"""
Run as Python to include matplotlib
***GTC script*** CTburden.py
Calculate uncertainty in CT ratio due to uncertainty in burden. Consider the problem
as the Molinger Gewecke circuit. Error = c/(c+s+b)-1 where c is the core impedance, s is the
secondary leakage impedance and b is the burden impedance.
Note that GTC trig functions do not work on reals/integers, we need to use the math library.
"""
from __future__ import division
import math
import cmath
import GTC
import numpy as np
import matplotlib.pyplot as plt

def error(core,secondary,burden):
    """calculates the error in the secondary current from impedance values"""
    return -(secondary+burden)/(core+secondary+burden)
    
def burdenVA(VA,pf,Isec):
    """calculates the complex impedance of a burden from its VA, PF and secondary current"""
    if pf >= 0:  #check sign convention for PF, '-'for capacitive
        multiplier = 1.0
    else:
        multiplier = -1.0
    magZ = VA/Isec**2
##     print magZ
##     print pf
##     pf = GTC.ureal(0.8,0.0001)
    phi = math.acos(pf)
    Z = magZ*math.cos(phi)+ multiplier*1j*magZ*math.sin(phi)
    return Z
    
def burdenZ(Z,Isec):
    """returns a burden in the form of VA and PF based on its impedance"""
    if Z.imag >= 0:  #check sign convention for PF, '-'for capacitive
        multiplier = 1.0
    else:
        multiplier = -1.0
    phi = cmath.phase(Z)
    PF = math.cos(phi)*multiplier
    magZ = abs(Z)
    VA = magZ*Isec**2
    return (VA, PF)
    
def coreZ(excite, frequency):
    """
    'excite' is the excitation level in %
    returns a complex impedance
    based on empirical fits
    """
    x = excite
    a1 = 11.7204144319303
    a2 = -0.0273771774408702
    a3 = 29.7364068940134
    inductance = (a1*math.log(x)+a2*x+a3)*1e-3
    
    a1 = 3.90272331971033
    a2 = 0.000243735060998679
    a3 = 5.8544055409592
    resistance = a1*math.log(x)+a2*x**2+a3/math.sqrt(x)
    return resistance + 1j*2*math.pi*frequency*inductance

VA = 5.0
isec = 5.0
z1 = burdenVA(VA,1.0,isec)
s = 0.022 +1j*0.01

x = [5,10,20,40,80,100,120]
errors = []
phase = []
for i in range(len(x)):
    coreimpedance = coreZ(x[i]*abs((s+z1)/0.2),50)#abs recognises that volts at core higher than at burden
    complex_error = error(coreimpedance,s,z1)
    print complex_error.real*100,complex_error.imag*100
    errors.append(complex_error.real*100)
    phase.append(complex_error.imag*100)

line1 = plt.plot(x,errors)
line2 = plt.plot(x,phase)
data1 = [-0.644,-0.569,-0.515,-0.472,-0.462,-0.441,-0.430]
line3 = plt.plot(x,data1,'r+')
data2 = [0.989,0.848,0.741,0.661,0.598,0.496,0.458]
line4 = plt.plot(x,data2,'gx')


errors2 = []
phase2 = []
#now shift to the 1.25 VA results
VA = 1.25
z2 = burdenVA(VA,1.0,isec)
## x2 = [] # excitation level for lower burden is divided by 4 to give correct c
## for xx in x:
##     x2.append(xx/1.0)
## print x2

for i in range(len(x)):
    coreimpedance = coreZ(x[i]*abs((s+z2)/0.2),50)#lower volts
    complex_error = error(coreimpedance,s,z2)
    print complex_error.real*100,complex_error.imag*100
    errors2.append(complex_error.real*100)
    phase2.append(complex_error.imag*100)
    
line5 = plt.plot(x,errors2)
line6 = plt.plot(x,phase2)
data1 = [-0.260,-0.211,-0.181,-0.161,-0.152,-0.142,-0.140]
line7 = plt.plot(x,data1,'yo')
data2 = [0.439,0.335,0.276,0.237,0.221,0.201,0.196]
line8 = plt.plot(x,data2,'bd')    
plt.show()


