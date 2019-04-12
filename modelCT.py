from __future__ import division
"""
The CT class optionally takes in either the real and imaginary part
of the secondary leakage impedance separately or as a single complex number,
in methods error or error_z. This was hurridly triggered by not being able to get
complex numbers into scipy optimize functions, but should be tidied up.
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
    Methods finishing in 'gtc' are using GTC rather than math functions. It is
    possible to conditionally select the function depending on type, but this
    is a future improvement after the likely use of this module is determined.
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
        
    def burdenVAgtc(self,VA,pf,Isec):
        """
        calculates the complex impedance of a burden from its VA, PF and secondary current
        """
        if pf >= 0:  #check sign convention for PF, '-'for capacitive
            multiplier = 1.0
        else:
            multiplier = -1.0
        magZ = VA/Isec**2
        phi = GTC.acos(pf)
        Z = magZ*GTC.cos(phi)+ multiplier*1j*magZ*GTC.sin(phi)
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

    def burdenZgtc(self,Z,Isec):
        """
        returns a burden in the form of VA and PF based on its impedance
        uses GTC functions to cope if Z is a GTC ucomplex
        """
        if Z.imag >= 0:  #check sign convention for PF, '-'for capacitive
            multiplier = 1.0
        else:
            multiplier = -1.0
        phase = GTC.phase(Z)
        # phi = cmath.phase(Z)
        PF = GTC.cos(phase)*multiplier
        magZ = GTC.magnitude(Z)
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
        
    def coreZgtc(self, excite, frequency):
        """
        'excite' is the excitation level in %
        returns a complex impedance
        based on empirical fits...this works wit ureal for x
        """
        x = excite
        inductance = (self.a1*GTC.log(x)+self.a2*x+self.a3)*1e-3
        resistance = self.a4*GTC.log(x)+self.a5*x**2+self.a6/GTC.sqrt(x)
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
        line3 = plt.plot(points,error_data,'r+')
        line4 = plt.plot(points,phase_data,'r+')
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
        
        correct_error = []#add corrections to the measured values
        correct_phase = []
        for i in range(len(points)):
            correct_error.append(error_data[i]+error_corrections[i])
            correct_phase.append(phase_data[i]+phase_corrections[i])
            
        return correct_error, correct_phase
    
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
    #should look at GTC output when model inputs have uncertainty from the fits
    a1 = GTC.ureal(11.7204144319303,0.346683597109428,34,label ='a1',dependent = True)
    a2 = GTC.ureal(-0.0273771774408702,0.0136282110753501,34,label ='a2',dependent = True)
    a3 = GTC.ureal(29.7364068940134,0.695696459888704,34,label ='a3',dependent = True)
    GTC.set_correlation(-0.826126693986887, a1, a2)
    GTC.set_correlation(-0.881067593505553, a1, a3)
    GTC.set_correlation(0.551078899199967, a2, a3)

    a4 = GTC.ureal(3.90272331971033,0.0859519946028212,34,label ='a4',dependent = True)
    a5 = GTC.ureal(0.000243735060998679,0.0000527749402101903,34,label ='a5',dependent = True)
    a6 = GTC.ureal(5.8544055409592,0.501146466659668,34,label ='a6',dependent = True)
    GTC.set_correlation(-0.679730943215757, a4, a5)
    GTC.set_correlation(-0.562016044432279, a4, a6)
    GTC.set_correlation(0.303221134084321, a5, a6)
    
    this_core_gtc = [a1,a2,a3,a4,a5,a6]
    this_CT_gtc = CT(this_core_gtc)
    print GTC.log(a1)
    x = GTC.ureal(100,1)
    answer = this_CT_gtc.coreZgtc(x, 50)
    error = repr(answer.real)
    print error
    burden = GTC.ucomplex((0.186 + 0.12j),(0.0001,0.0001))
    print(this_CT.burdenZgtc(burden, 5.0))