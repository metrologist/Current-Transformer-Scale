"""
Run as Python to include matplotlib
CT_burden_analyser.py is using data from S19821 1 turn V5.0_draft_KJ.xls
Tidying up program structure
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
        
    def error(self,core,secondary,burden):
        """
        calculates the error in the secondary current from impedance values
        of the core, secondary and burden
        """
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
        return math.sqrt(residual_sum)
        
    def table_analysis(self, sec, burdn, points, error_data, phase_data):
        errors = []
        phase = []
        for i in range(len(points)):
            coreimpedance = self.coreZ(points[i]*abs((sec+burdn)/0.2),50)#abs recognises that volts at core higher than at burden
            complex_error = self.error(coreimpedance,sec,burdn)
            errors.append(complex_error.real*100)
            phase.append(complex_error.imag*100)
        s1 = self.fit_check(errors, error_data)
        s2 = self.fit_check(phase, phase_data)
        return s1 + s2
        
    def plot_table_analysis(self, sec, burdn, points, error_data, phase_data):
        errors = []
        phase = []
        for i in range(len(points)):
            coreimpedance = self.coreZ(points[i]*abs((sec+burdn)/0.2),50)#abs recognises that volts at core higher than at burden
            complex_error = self.error(coreimpedance,sec,burdn)
            errors.append(complex_error.real*100)
            phase.append(complex_error.imag*100)
        line1 = plt.plot(points,errors)
        line2 = plt.plot(points,phase)
        line3 = plt.plot(x,error_data,'r+')
        line4 = plt.plot(x,phase_data,'r+')
        plt.show()

    def burden_correction(self, nominal_burdn, sec, burdn, points, error_data, phase_data):
        """
        First calculates the errors at the exact nominal burden and the actual burden
        and then uses the differences between these two as a correction to the measured results.
        """
        errors = []#using the measured value of burden
        phase = []
        for i in range(len(points)):
            coreimpedance = self.coreZ(points[i]*abs((sec+burdn)/0.2),50)#abs recognises that volts at core higher than at burden
            complex_error = self.error(coreimpedance,sec,burdn)
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
    
    def error_vs_burden(burden_list, sec, excitation):
        """
        Calculates error vs burden to check if spreadsheet extrapolation methods are valid
        """
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
    this_core = [11.7204144319303, -0.0273771774408702, 29.7364068940134, 3.90272331971033, 0.000243735060998679, 5.8544055409592]
    this_CT = CT(this_core)
##     s = 0.024286886*1.4 +1j*0.001*1.0 #critical secondary leakage impedance for the CT, ideally should be in the CT object
    s = 0.02657786 + 1j*0.00029099 #from find_sec.py
    
##     x = [5,10,20,40,60,100,120]
##     nominal_burden = this_CT.burdenVA(5,0.8,5)
##     print 'nominal burden =', nominal_burden
##     zH = 0.168 +1j*0.1251
##     
##     print '5 VA 0.8 PF from MSL inductive ~5 VA 0.8 PF'
##     this_CT.burden_correction(nominal_burden,s, zH, x,[-1.08, -0.946, -0.844, -0.766, -0.722, -0.656, -0.624],[0.509, 0.436, 0.375, 0.332, 0.285, 0.201, 0.175])
    
    #estimate corrections for pt participants not at correct nominal burdens
##     xpt = [5,20,100,120]
##     nominal_burden = this_CT.burdenVA(5,0.8,5)
##     print 'nominal burden =', nominal_burden
##     broad_spectrum_bdn = this_CT.burdenVA(5,1.0,5)
##     print 'broad spectrum 5 VA =', broad_spectrum_bdn
##     this_CT.burden_correction(nominal_burden,s, broad_spectrum_bdn, xpt,[0,0,0,0],[0,0,0,0])#note putting in '0' data as correction will be applied in excel

##     nominal_burden = this_CT.burdenVA(1.25,1.0,5)
##     print 'nominal burden =', nominal_burden
##     electrix_bdn = this_CT.burdenVA(1.2,1.0,5)
##     print 'electrix 1.25 VA =', electrix_bdn
##     this_CT.burden_correction(nominal_burden,s, electrix_bdn, xpt,[0,0,0,0],[0,0,0,0])#note putting in '0' data as correction will be applied in excel


##     nominal_burden = this_CT.burdenVA(1.25,1.0,5)
##     print 'nominal burden =', nominal_burden
##     bs_bdn = this_CT.burdenVA(1.123,1.0,5)
##     print 'broad_spectrum 1.25 VA =', bs_bdn
##     this_CT.burden_correction(nominal_burden,s, bs_bdn, xpt,[0,0,0,0],[0,0,0,0])#note putting in '0' data as correction will be applied in excel

    #plot tables
    zA = 0.2580 + 1j*0.0058
    zB = 0.2170 + 1j*0.0029 #5VA burden B
    zF = 0.0207 + 1j*0.0008
    zC = 0.1800 + 1j*0.0029
    zE = 0.0605 + 1j*0.0022
    zD = 0.1010 + 1j*0.0
    zH = 0.1680 + 1j*0.1251
    x = [5,10,20,40,60,100,120]
    
##     this_CT.plot_table_analysis(s, zA, x,[-0.776,-0.690,-0.624,-0.578,-0.562,-0.536,-0.520],[1.169,1.004,0.890,0.785,0.704,0.564,0.518])
##     this_CT.plot_table_analysis(s, zB, x,[-0.682,-0.604,-0.548,-0.502,-0.492,-0.470,-0.458],[1.041,0.896,0.785,0.701,0.634,0.524,0.483])
##     this_CT.plot_table_analysis(s, zF, x,[-0.224,-0.185,-0.155,-0.138,-0.130,-0.121,-0.119],[0.393,0.300,0.239,0.205,0.189,0.176,0.172])
##     this_CT.plot_table_analysis(s, zC, x,[-0.600,-0.522,-0.474,-0.434,-0.418,-0.404,-0.398],[0.925,0.785,0.684,0.608,0.559,0.477,0.442])
##     this_CT.plot_table_analysis(s, zE, x,[-0.326,-0.273,-0.240,-0.217,-0.204,-0.191,-0.190],[0.535,0.425,0.361,0.311,0.291,0.265,0.259])
##     this_CT.plot_table_analysis(s, zD, x,[-0.420,-0.360,-0.320,-0.291,-0.276,-0.264,-0.261],[0.684,0.547,0.471,0.416,0.390,0.352,0.335])
    this_CT.plot_table_analysis(s, zH, x,[-1.040,-0.944,-0.840,-0.762,-0.722,-0.652,-0.622],[0.521,0.439,0.378,0.326,0.285,0.201,0.175])

    
## VA = 5.0
## isec = 5.0


## # start with s = 0.022 +1j*0.01
## aa =0.1
## s = 0.024286886*1.4 +1j*0.001*1.0


## bd_list = [0.0,0.1,0.2,0.3,0.4,0.5,1,1.5,2,3,4,5]
## my_list = []
## for i in range(len(bd_list)):
##     my_list.append(burdenVA(bd_list[i],1.0,5))

## error_vs_burden(my_list,s,100)


## x = [5,10,20,40,60,100,120]

## zB = 0.2170*1.0 +1j*0.0029 #5VA burden B
## sum1 = table_analysis(s, zB, x,[-0.682,-0.604,-0.548,-0.502,-0.492,-0.470,-0.458],[1.041,0.896,0.785,0.701,0.634,0.524,0.483])
## print sum1

## zH = 0.168 +1j*0.1251
## sum2 = table_analysis(s, zH, x,[-1.08, -0.946, -0.844, -0.766, -0.722, -0.656, -0.624],[0.509, 0.436, 0.375, 0.332, 0.285, 0.201, 0.175])
## print sum2

## zF = 0.0207 + 1j*0.0008
## sum3 = table_analysis(s, zF, x,[-0.224,-0.185,-0.155,-0.138,-0.130,-0.121,-0.119],[0.393, 0.300, 0.239, 0.205, 0.189, 0.176, 0.172])
## print sum3

## zA = 0.2580 + 1j*0.0058
## sum4 = table_analysis(s, zA, x,[-0.776,-0.690,-0.624,-0.578,-0.562,-0.536,-0.520],[1.169,1.004,0.890,0.785,0.704,0.564,0.518])
## print sum4

## zC = 0.1800 +1j*0.0029
## sum5 = table_analysis(s, zC, x,[-0.600,-0.522,-0.474,-0.434,-0.418,-0.404,-0.398],[0.925,0.785,0.684,0.608,0.559,0.477,0.442])
## print sum5

## zE = 0.0605 +1j*0.0022
## sum6 = table_analysis(s, zE, x,[-0.326,-0.273,-0.240,-0.217,-0.204,-0.191,-0.190],[0.535,0.425,0.361,0.311,0.291,0.265,0.259])
## print sum6

## print (sum1 + sum2 +sum3 + sum4 + sum5 +sum6)/6.0
## plot_table_analysis(s, zE, x,[-0.326,-0.273,-0.240,-0.217,-0.204,-0.191,-0.190],[0.535,0.425,0.361,0.311,0.291,0.265,0.259])


## nominal_burden = burdenVA(5,0.8,5)
## print nominal_burden
## print '5 VA 0.8 PF from MSL inductive ~5 VA 0.8 PF'
## burden_correction(nominal_burden,s, zH, x,[-1.08, -0.946, -0.844, -0.766, -0.722, -0.656, -0.624],[0.509, 0.436, 0.375, 0.332, 0.285, 0.201, 0.175])
## print '5 VA 0.8 PF from MSL non-inductive ~5 VA 1.0 PF'
## burden_correction(nominal_burden,s, zB, x, [-0.694,-0.616,-0.558,-0.508,-0.499,-0.478,-0.462],[1.053,0.908,0.797,0.713,0.643,0.527,0.483])
## nominal_burden = burdenVA(1.25,1.0,5)
## print '1.25 VA 1.0 PF from MSL non-inductive 1 VA 1.0 PF'
## print nominal_burden
## burden_correction(nominal_burden,s, zE, x, [-0.329,-0.278,-0.240,-0.218,-0.206,-0.193,-0.190],[0.544,0.436,0.364,0.317,0.294,0.271,0.262])

## print '10 August 5VA 0.8 PF'
## nominal_burden = burdenVA(5,0.8,5)
## print nominal_burden
## print '5 VA 0.8 PF from MSL inductive ~5 VA 0.8 PF'
## burden_correction(nominal_burden,s, zH, x,[-1.040,-0.944,-0.840,-0.762,-0.722,-0.652,-0.622],[0.521,0.439,0.378,0.326,0.285,0.201,0.175])    

##     print '5 VA 0.8 PF from MSL non-inductive ~5 VA 1.0 PF'
##     burden_correction(nominal_burden,s, zB, x, [-0.694,-0.616,-0.558,-0.508,-0.499,-0.478,-0.462],[1.053,0.908,0.797,0.713,0.643,0.527,0.483])
##     nominal_burden = burdenVA(1.25,1.0,5)
##     print '1.25 VA 1.0 PF from MSL non-inductive 1 VA 1.0 PF'
##     print nominal_burden
##     burden_correction(nominal_burden,s, zE, x, [-0.329,-0.278,-0.240,-0.218,-0.206,-0.193,-0.190],[0.544,0.436,0.364,0.317,0.294,0.271,0.262])

##     print '10 August 5VA 0.8 PF'
##     nominal_burden = burdenVA(5,0.8,5)
##     print nominal_burden
##     print '5 VA 0.8 PF from MSL inductive ~5 VA 0.8 PF'
##     burden_correction(nominal_burden,s, zH, x,[-1.040,-0.944,-0.840,-0.762,-0.722,-0.652,-0.622],[0.521,0.439,0.378,0.326,0.285,0.201,0.175])    
    
    