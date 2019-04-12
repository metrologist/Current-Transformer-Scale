from __future__ import division
"""
Use least squares to choose best value of secondary leakage impedance
The CT class has been modified here to take in the real and imaginary part
of the secondary leakage separately, rather than as a complex number.
"""
from scipy.optimize import minimize

import math
import cmath
import GTC
import numpy as np
import matplotlib.pyplot as plt
import ExcelPython

class CT(object):
    """
    Circuit model of a current transformer with methods to build the model
    from measurement data and methods to then calculate the error and phase
    displacement at different currents and burdens.
    """
    def __init__(self, core):
        """
        'core' is a list of equation coefficients that describe the magnetising
        impedance of the core.
        """
        self.a1 = core[0]#all used in coreZ below
        self.a2 = core[1]
        self.a3 = core[2]
        self.a4 = core[3]
        self.a5 = core[4]
        self.a6 = core[5]
        
    def error(self,core,r,l,burden):
        """
        calculates the error in the secondary current from impedance values
        of the core, secondary and burden
        """
        secondary = r + 1j*l
        return -(secondary+burden)/(core+secondary+burden)

    def error_z(self,core,z,burden):
        """
        calculates the error in the secondary current from impedance values
        of the core, secondary and burden. z is a ucomplex
        """
        secondary = z
        return -(secondary+burden)/(core+secondary+burden)

    def burdenVA(self,VA,pf,Isec):
        """
        calculates the complex impedance of a burden from its VA, PF and secondary current
        """
        if pf >= 0:  #check sign convention for PF, '-'for capacitive
            multiplier = 1.0
        else:
            multiplier = -1.0
        magZ = VA/Isec**2
        phi = math.acos(pf)
        Z = magZ*math.cos(phi)+ multiplier*1j*magZ*math.sin(phi)
        return Z
        
    def burdenZ(self,Z,Isec):
        """
        returns a burden in the form of VA and PF based on its impedance
        """
        if Z.imag >= 0:  #check sign convention for PF, '-'for capacitive
            multiplier = 1.0
        else:
            multiplier = -1.0
        phi = cmath.phase(Z)
        PF = math.cos(phi)*multiplier
        magZ = abs(Z)
        VA = magZ*Isec**2
        return (VA, PF)
        
    def coreZ(self, excite, frequency):
        """
        'excite' is the excitation level in %
        returns a complex impedance
        based on empirical fits
        """
        x = excite
        inductance = (self.a1*math.log(x)+self.a2*x+self.a3)*1e-3
        resistance = self.a4*math.log(x)+self.a5*x**2+self.a6/math.sqrt(x)
        return resistance + 1j*2*math.pi*frequency*inductance
        
    def fit_check(self,lista, listb):
        residual_sum = 0
        assert len(lista)==len(listb),"lists are different lengths!"
        for i in range(len(lista)):
            residual_sum = residual_sum + (lista[i]-listb[i])**2
        return residual_sum
        
    def table_analysis(self, r, l, burdn, points, error_data, phase_data):
        sec =r + 1j*l
        errors = []
        phase = []
        for i in range(len(points)):
            coreimpedance = self.coreZ(points[i]*abs((sec+burdn)/0.2),50)#abs recognises that volts at core higher than at burden
            complex_error = self.error(coreimpedance,r,l,burdn)
            errors.append(complex_error.real*100)
            phase.append(complex_error.imag*100)
        s1 = self.fit_check(errors, error_data)
        s2 = self.fit_check(phase, phase_data)
        return s1 + s2
        
    def plot_table_analysis(self, r, l, burdn, points, error_data, phase_data,title):
        plt.close()#useful for when repeat runs are used
        plt.title(title)
        sec = r +1j*l
        errors = []
        phase = []
        for i in range(len(points)):
            coreimpedance = self.coreZ(points[i]*abs((sec+burdn)/0.2),50)#abs recognises that volts at core higher than at burden
            complex_error = self.error(coreimpedance,r,l,burdn)
            errors.append(complex_error.real*100)
            phase.append(complex_error.imag*100)
        line1 = plt.plot(points,errors)
        line2 = plt.plot(points,phase)
        line3 = plt.plot(x,error_data,'r+')
        line4 = plt.plot(x,phase_data,'r+')
        plt.show()

    def burden_correction(self, nominal_burdn, r, l, burdn, points, error_data, phase_data):
        """
        First calculates the errors at the exact nominal burden and the actual burden
        and then uses the differences between these two as a correction to the measured results.
        """
        sec = r + 1j*l
        errors = []#using the measured value of burden
        phase = []
        for i in range(len(points)):
            coreimpedance = self.coreZ(points[i]*abs((sec+burdn)/0.2),50)#abs recognises that volts at core higher than at burden
            complex_error = self.error(coreimpedance,r,l,burdn)
            errors.append(complex_error.real*100)
            phase.append(complex_error.imag*100)
            
        errors_nom = []#using the nominal value of burden
        phase_nom = []
        for i in range(len(points)):
            coreimpedance = self.coreZ(points[i]*abs((sec+nominal_burdn)/0.2),50)#abs recognises that volts at core higher than at burden
            complex_error = self.error(coreimpedance,sec,nominal_burdn)
            errors_nom.append(complex_error.real*100)
            phase_nom.append(complex_error.imag*100)
            
        error_corrections = []
        phase_corrections = []
        for i in range(len(points)):
            error_corrections.append(errors_nom[i]-errors[i])
            phase_corrections.append(phase_nom[i]-phase[i])
            
        print error_corrections
        print phase_corrections
        
        correct_error = []#add corrections to the measured values
        correct_phase = []
        for i in range(len(points)):
            correct_error.append(error_data[i]+error_corrections[i])
            correct_phase.append(phase_data[i]+phase_corrections[i])
            
        print correct_error
        print correct_phase
        
    def burden_correction_z(self, nominal_burdn, z, burdn, points, error_data, phase_data):
        """
        First calculates the errors at the exact nominal burden and the actual burden
        and then uses the differences between these two as a correction to the measured results.
        """
##         sec = r + 1j*l
        sec = z
        errors = []#using the measured value of burden
        phase = []
        for i in range(len(points)):
            coreimpedance = self.coreZ(points[i]*abs((sec+burdn)/0.2),50)#abs recognises that volts at core higher than at burden
            complex_error = self.error_z(coreimpedance,z,burdn)
            errors.append(complex_error.real*100)
            phase.append(complex_error.imag*100)
            
        errors_nom = []#using the nominal value of burden
        phase_nom = []
        for i in range(len(points)):
            coreimpedance = self.coreZ(points[i]*abs((sec+nominal_burdn)/0.2),50)#abs recognises that volts at core higher than at burden
            complex_error = self.error_z(coreimpedance,sec,nominal_burdn)
            errors_nom.append(complex_error.real*100)
            phase_nom.append(complex_error.imag*100)
            
        error_corrections = []
        phase_corrections = []
        for i in range(len(points)):
            error_corrections.append(errors_nom[i]-errors[i])
            phase_corrections.append(phase_nom[i]-phase[i])
            
        print error_corrections
        print phase_corrections
        
        correct_error = []#add corrections to the measured values
        correct_phase = []
        for i in range(len(points)):
            correct_error.append(error_data[i]+error_corrections[i])
            correct_phase.append(phase_data[i]+phase_corrections[i])
            
        print correct_error
        print correct_phase
    
    def error_vs_burden(burden_list, r, l, excitation):
        """
        Calculates error vs burden to check if spreadsheet extrapolation methods are valid
        """
        sec = r +1j*l
        x_axis = []
        this_error = []
        this_phase = []
        print burden_list
        for i in range(len(burden_list)):
            coreimpedance = self.coreZ(excitation*abs((sec + burden_list[i])/0.2),50)
            complex_error = self.error(coreimpedance,sec,burden_list[i])
            x_axis.append(burden_list[i].real)
            this_error.append(complex_error.real*100)
            this_phase.append(complex_error.imag*100)
        print this_error
        print this_phase
        line1 = plt.plot(x_axis,this_error)
        line2 = plt.plot(x_axis,this_phase)
        plt.show()



if __name__ == "__main__":
    #use fitted impedance coefficients to define behviour of core impedance
    this_core = [11.7204144319303, -0.0273771774408702, 29.7364068940134, 3.90272331971033, 0.000243735060998679, 5.8544055409592]
    this_CT = CT(this_core)
    	
    x = [5,10,20,40,60,100,120]
    #data will be taken from the ForPython worksheet
    ianz = ExcelPython.CALCULATOR("S21982 1 turn V5.0_draft_KJ.xlsm", "myResults.xlsx")
    #copy a block from the source worksheet
    block_descriptor = [1,343,1,8]
    my_copy_data = ianz.getdata_block('ForPython', block_descriptor)
    #essentially a hand selection process
    a1 = ianz.extract_column(my_copy_data, 6, [40,46])
    b1 = ianz.extract_column(my_copy_data, 7, [40,46])
    
    a2 = ianz.extract_column(my_copy_data, 6, [125,131])
    b2 = ianz.extract_column(my_copy_data, 7, [125,131])
    
    a3 = ianz.extract_column(my_copy_data, 6, [133,139])
    b3 = ianz.extract_column(my_copy_data, 7, [133,139])
    
    a4 = ianz.extract_column(my_copy_data, 6, [57,63])
    b4 = ianz.extract_column(my_copy_data, 7, [57,63])
    
    a5 = ianz.extract_column(my_copy_data, 6, [91,97])
    b5 = ianz.extract_column(my_copy_data, 7, [91,97])
    
    a6 = ianz.extract_column(my_copy_data, 6, [74,80])
    b6 = ianz.extract_column(my_copy_data, 7, [74,80])
    
    a7 = ianz.extract_column(my_copy_data, 6, [159,165])
    b7 = ianz.extract_column(my_copy_data, 7, [159,165])
    
    a8 = ianz.extract_column(my_copy_data, 6, [210,216])
    b8 = ianz.extract_column(my_copy_data, 7, [210,216])
    
    a9 = ianz.extract_column(my_copy_data, 6, [227,233])
    b9 = ianz.extract_column(my_copy_data, 7, [227,233])
    
    a10 = ianz.extract_column(my_copy_data, 6, [184,190])
    b10 = ianz.extract_column(my_copy_data, 7, [184,190])
    
    zA = 0.2580 + 1j*0.0058
    zB = 0.2170 + 1j*0.0029 #5VA burden B
    zF = 0.0207 + 1j*0.0008
    zC = 0.1800 + 1j*0.0029
    zE = 0.0605 + 1j*0.0022
    zD = 0.1010 + 1j*0.0
    zH = 0.1680 + 1j*0.1251
    
    def fitfun(a):#essential that the selected burden matches the data!
        sum1 = this_CT.table_analysis(a[0], a[1], zA, x,a1,b1)
        sum2 = this_CT.table_analysis(a[0], a[1], zB, x,a2,b2)
        sum3 = this_CT.table_analysis(a[0], a[1], zF, x,a3,b3)
        sum4 = this_CT.table_analysis(a[0], a[1], zC, x,a4,b4)
        sum5 = this_CT.table_analysis(a[0], a[1], zE, x,a5,b5)
        sum6 = this_CT.table_analysis(a[0], a[1], zD, x,a6,b6)
        sum7 = this_CT.table_analysis(a[0], a[1], zH, x,a7,b7)
        sum8 = this_CT.table_analysis(a[0], a[1], zB, x,a8,b8)
        sum9 = this_CT.table_analysis(a[0], a[1], zH, x,a9,b9)
        sum10 = this_CT.table_analysis(a[0], a[1], zF, x,a10,b10)
        total = sum1 + sum2 + sum3 + sum4 + sum5 + sum6 + sum7 + sum8 + sum9 + sum10
        return total
        
    result = minimize(fitfun,[0.05,0.001])#, method = 'Nelder-Mead')#need to start with good guess
    print result
    #Make it clearer where the data comes from    
    

##     print my_copy_data
    #put the block, unchanged into the output spreadsheet
##     ianz.makeworkbook(my_copy_data, 'fromS21982_1_turn')

a = result.x
#close a plot to see the next one!
## this_CT.plot_table_analysis(a[0], a[1], zA, x,a1,b1,'6VA 1.0')
## this_CT.plot_table_analysis(a[0], a[1], zB, x,a2,b2,'5VA 1.0')
## this_CT.plot_table_analysis(a[0], a[1], zF, x,a3,b3,'0VA')
## this_CT.plot_table_analysis(a[0], a[1], zC, x,a4,b4,'4VA 1.0')
## this_CT.plot_table_analysis(a[0], a[1], zE, x,a5,b5,'1VA 1.0')
## this_CT.plot_table_analysis(a[0], a[1], zD, x,a6,b6,'2VA 1.0')
## this_CT.plot_table_analysis(a[0], a[1], zH, x,a7,b7,'5VA 0.8')
## this_CT.plot_table_analysis(a[0], a[1], zB, x,a8,b8,'5VA 1.0')
## this_CT.plot_table_analysis(a[0], a[1], zH, x,a9,b9,'5VA 0.8')
## this_CT.plot_table_analysis(a[0], a[1], zF, x,a10,b10,'0VA')

hessian = result.hess_inv #inverse Hessian is estimate of variance/covariance matrix
sum_square = result.fun#final rss

#scaling the inverse hessian by multiplying by the residuals
sa0 = math.sqrt(hessian[0,0]*sum_square)
sa1 = math.sqrt(hessian[1,1]*sum_square)

#check that adding a standard deviation roughly doubles the residual sum square

chi_square = fitfun(a)
print 'best fit=', chi_square
bb =[a[0]+sa0, a[1]]
print 'with r sd added =', fitfun(bb)/chi_square
cc = [a[0],a[1]+sa1]
print 'wih l sd added =', fitfun(cc)/chi_square
print 't-test', a[0]/sa0,a[1]/sa1
dd = [a[0]+sa0,a[1]+sa1]
print 'wih r and l sd added =', fitfun(dd)/chi_square

#now use GTC with ucomplex

zs = GTC.ucomplex(a[0]+1j*a[1],(hessian[0,0]*sum_square,hessian[0,1]*sum_square,hessian[1,0]*sum_square,hessian[1,1]*sum_square),138)
## zs = GTC.ucomplex(a[0]+1j*a[1],(math.sqrt(hessian[0,0]*sum_square),math.sqrt(hessian[1,1]*sum_square)),138)

excite = 60
core = this_CT.coreZ(excite,50)
print
print 'calculate error with zH'
complex_error = this_CT.error_z(core, zs, zH)*100
print repr(complex_error)
print complex_error.u
print
print 'calculate error with zB'
complex_error2 = this_CT.error_z(core, zs, zB)*100
print repr(complex_error2)
print complex_error2.u
print
print'calculate difference'
correction = complex_error - complex_error2
print repr(correction)
print correction.u

target_burden = this_CT.burdenVA(5.0,0.8,5.0)
#using 'complex(zs)' breaks the GTC chain, which might be needed
this_CT.burden_correction_z(target_burden,complex(zs),zH,x,a9,b9)


"""
>pythonw -u "find_sec.py"
   status: 0
  success: True
     njev: 12
     nfev: 48
 hess_inv: array([[  6.12700293e-04,  -1.24604686e-06],
       [ -1.24604686e-06,   3.82743402e-04]])
      fun: 0.05632551657433842
        x: array([  2.68799510e-02,  -8.07464912e-05])
  message: 'Optimization terminated successfully.'
      jac: array([  7.89295882e-07,   1.09011307e-06])
best fit= 0.0563255165743
with r sd added = 1.49337576042
wih l sd added = 1.52500788018
t-test 4.57563988929 -0.0173907059365
wih r and l sd added = 1.99613631484

calculate error with zH
ucomplex((-0.72386346771923082+0.27461189822665366j), u=[0.016964935823619881,0.018167488004654963], r=-0.221, df=138)
standard_uncertainty(real=0.016964935823619878, imag=0.018167488004654963)

calculate error with zB
ucomplex((-0.49672399382186078+0.64918433823687893j), u=[0.017073634505796072,0.018222932200190251], r=-0.222, df=138)
standard_uncertainty(real=0.017073634505796072, imag=0.018222932200190247)

calculate difference
ucomplex((-0.2271394738973700478-0.3745724400102252716j), u=[0.0001720422965400067873,0.000136501711529169893], r=0.04200000000000000261, df=138)
standard_uncertainty(real=0.0001720422965400068, imag=0.00013650171152916984)
[0.034973061015349804, 0.03139963919028943, 0.028414347363219683, 0.02579285242517615, 0.024157764675693527, 0.02111530552223917, 0.019367132398266818]
[-0.015881153041213514, -0.01314214569775779, -0.01080296754187271, -0.008154906556202657, -0.005915684384622366, -0.0014594280122681436, 0.0007053121766522341]
[-1.0550269389846503, -0.9206003608097105, -0.8195856526367803, -0.7402071475748239, -0.6998422353243064, -0.6348846944777609, -0.6046328676017332]
[0.5106265046437426, 0.42900793147413896, 0.3702605858102226, 0.3263665334093772, 0.2849725242810992, 0.2050712001403942, 0.1781471194627424]
>Exit code: 0
"""

