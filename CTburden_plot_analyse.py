"""
Run as Python to include matplotlib
CT_burden_plot_analyse.py is using data from S19821 1 turn V5.0_draft_KJ.xls
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
    
def fit_check(lista, listb):
    residual_sum = 0
    assert len(lista)==len(listb),"lists are different lengths!"
    for i in range(len(lista)):
        residual_sum = residual_sum + (lista[i]-listb[i])**2
    return math.sqrt(residual_sum)
    
def table_analysis(sec, burdn, points, error_data, phase_data):
    errors = []
    phase = []
    for i in range(len(points)):
        coreimpedance = coreZ(points[i]*abs((sec+burdn)/0.2),50)#abs recognises that volts at core higher than at burden
        complex_error = error(coreimpedance,sec,burdn)
        errors.append(complex_error.real*100)
        phase.append(complex_error.imag*100)
        
    s1 = fit_check(errors, error_data)
    s2 = fit_check(phase, phase_data)
    return s1 + s2
        
def plot_table_analysis(sec, burdn, points, error_data, phase_data):
    errors = []
    phase = []
    for i in range(len(points)):
        coreimpedance = coreZ(points[i]*abs((sec+burdn)/0.2),50)#abs recognises that volts at core higher than at burden
        complex_error = error(coreimpedance,sec,burdn)
        errors.append(complex_error.real*100)
        phase.append(complex_error.imag*100)
    line1 = plt.plot(points,errors)
    line2 = plt.plot(points,phase)
    line3 = plt.plot(x,error_data,'r+')
    line4 = plt.plot(x,phase_data,'r+')
    plt.show()

    return 

def burden_correction(nominal_burdn, sec, burdn, points, error_data, phase_data):
    """
    First calculates the errors at the exact nominal burden and the actual burden
    and then uses the differences between these two as a correction to the measured results.
    """
    errors = []#using the measured value of burden
    phase = []
    for i in range(len(points)):
        coreimpedance = coreZ(points[i]*abs((sec+burdn)/0.2),50)#abs recognises that volts at core higher than at burden
        complex_error = error(coreimpedance,sec,burdn)
        errors.append(complex_error.real*100)
        phase.append(complex_error.imag*100)
        
    errors_nom = []#using the nominal value of burden
    phase_nom = []
    for i in range(len(points)):
        coreimpedance = coreZ(x[i]*abs((sec+nominal_burdn)/0.2),50)#abs recognises that volts at core higher than at burden
        complex_error = error(coreimpedance,sec,nominal_burdn)
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
        coreimpedance = coreZ(excitation*abs((sec + burden_list[i])/0.2),50)
        complex_error = error(coreimpedance,sec,burden_list[i])
        x_axis.append(burden_list[i].real)
        this_error.append(complex_error.real*100)
        this_phase.append(complex_error.imag*100)
    print this_error
    print this_phase
    line1 = plt.plot(x_axis,this_error)
    line2 = plt.plot(x_axis,this_phase)
    plt.show()

VA = 5.0
isec = 5.0
## z1 = burdenVA(VA,1.0,isec)

# start with s = 0.022 +1j*0.01
aa =0.1
s = 0.024286886*1.4 +1j*0.001*1.0


## bd_list = [0.0,0.1,0.2,0.3,0.4,0.5,1,1.5,2,3,4,5]
## my_list = []
## for i in range(len(bd_list)):
##     my_list.append(burdenVA(bd_list[i],1.0,5))

## error_vs_burden(my_list,s,100)


x = [5,10,20,40,60,100,120]

zB = 0.2170*1.0 +1j*0.0029 #5VA burden B
sum1 = table_analysis(s, zB, x,[-0.682,-0.604,-0.548,-0.502,-0.492,-0.470,-0.458],[1.041,0.896,0.785,0.701,0.634,0.524,0.483])
print sum1

zH = 0.168 +1j*0.1251
sum2 = table_analysis(s, zH, x,[-1.08, -0.946, -0.844, -0.766, -0.722, -0.656, -0.624],[0.509, 0.436, 0.375, 0.332, 0.285, 0.201, 0.175])
print sum2

zF = 0.0207 + 1j*0.0008
sum3 = table_analysis(s, zF, x,[-0.224,-0.185,-0.155,-0.138,-0.130,-0.121,-0.119],[0.393, 0.300, 0.239, 0.205, 0.189, 0.176, 0.172])
print sum3

zA = 0.2580 + 1j*0.0058
sum4 = table_analysis(s, zA, x,[-0.776,-0.690,-0.624,-0.578,-0.562,-0.536,-0.520],[1.169,1.004,0.890,0.785,0.704,0.564,0.518])
print sum4

zC = 0.1800 +1j*0.0029
sum5 = table_analysis(s, zC, x,[-0.600,-0.522,-0.474,-0.434,-0.418,-0.404,-0.398],[0.925,0.785,0.684,0.608,0.559,0.477,0.442])
print sum5

zE = 0.0605 +1j*0.0022
sum6 = table_analysis(s, zE, x,[-0.326,-0.273,-0.240,-0.217,-0.204,-0.191,-0.190],[0.535,0.425,0.361,0.311,0.291,0.265,0.259])
print sum6

print (sum1 + sum2 +sum3 + sum4 + sum5 +sum6)/6.0
plot_table_analysis(s, zE, x,[-0.326,-0.273,-0.240,-0.217,-0.204,-0.191,-0.190],[0.535,0.425,0.361,0.311,0.291,0.265,0.259])

"""
>pythonw -u "CTburden_plot_analyse.py"
0.166873023323
0.159344501422
0.0508408425224
0.192004771656
0.165002747068
0.0841267558193
0.136365440302
>Exit code: 0
"""
"""
The above rss terms show that the 6 core impedance function coefficients together with the resistance and inductance of the
leakage impedance,s, agree reasonably well with measurements made across a range of burdens.
Next we need to use this model to take measurements made with a close to nominal burden and then correct these to give the
errors at the exact nominal burden.
"""
nominal_burden = burdenVA(5,0.8,5)
print nominal_burden
print '5 VA 0.8 PF from MSL inductive ~5 VA 0.8 PF'
burden_correction(nominal_burden,s, zH, x,[-1.08, -0.946, -0.844, -0.766, -0.722, -0.656, -0.624],[0.509, 0.436, 0.375, 0.332, 0.285, 0.201, 0.175])
print '5 VA 0.8 PF from MSL non-inductive ~5 VA 1.0 PF'
burden_correction(nominal_burden,s, zB, x, [-0.694,-0.616,-0.558,-0.508,-0.499,-0.478,-0.462],[1.053,0.908,0.797,0.713,0.643,0.527,0.483])
nominal_burden = burdenVA(1.25,1.0,5)
print '1.25 VA 1.0 PF from MSL non-inductive 1 VA 1.0 PF'
print nominal_burden
burden_correction(nominal_burden,s, zE, x, [-0.329,-0.278,-0.240,-0.218,-0.206,-0.193,-0.190],[0.544,0.436,0.364,0.317,0.294,0.271,0.262])

print '10 August 5VA 0.8 PF'
nominal_burden = burdenVA(5,0.8,5)
print nominal_burden
print '5 VA 0.8 PF from MSL inductive ~5 VA 0.8 PF'
burden_correction(nominal_burden,s, zH, x,[-1.040,-0.944,-0.840,-0.762,-0.722,-0.652,-0.622],[0.521,0.439,0.378,0.326,0.285,0.201,0.175])

"""
>pythonw -u "CTburden_plot_analyse.py"
0.166873023323
0.159344501422
0.0508408425224
0.192004771656
0.165002747068
0.0841267558193
0.136365440302
(0.16+0.12j)
5 VA 0.8 PF from MSL inductive ~5 VA 0.8 PF
[0.034889839132444145, 0.031326736796737364, 0.028351131528221774, 0.025732251354775304, 0.0240820337462897, 0.020950345303942908, 0.019129713985323482]
[-0.015591868817029342, -0.012921212568458296, -0.010613537820297447, -0.007947037281609504, -0.005653852295755324, -0.0010655619220558965, 0.001150923960077349]
[-1.045110160867556, -0.9146732632032626, -0.8156488684717782, -0.7402677486452247, -0.6979179662537103, -0.6350496546960571, -0.6048702860146765]
[0.49340813118297067, 0.4230787874315417, 0.36438646217970255, 0.3240529627183905, 0.27934614770424465, 0.19993443807794412, 0.17615092396007734]
5 VA 0.8 PF from MSL non-inductive ~5 VA 1.0 PF
[-0.3908460198837027, -0.3221146462497734, -0.2674275910508359, -0.22022162304010062, -0.19204383603151576, -0.14995117813020098, -0.1319970557730143]
[-0.5635418475575209, -0.48905141115014683, -0.43236468226101654, -0.3865026055287343, -0.3617228938158925, -0.32710977133297814, -0.3118624111837379]
[-1.0848460198837027, -0.9381146462497734, -0.825427591050836, -0.7282216230401006, -0.6910438360315158, -0.627951178130201, -0.5939970557730143]
[0.48945815244247903, 0.4189485888498532, 0.3646353177389835, 0.3264973944712657, 0.2812771061841075, 0.19989022866702189, 0.17113758881626207]
1.25 VA 1.0 PF from MSL non-inductive 1 VA 1.0 PF
(0.05+0j)
[0.03864785179957281, 0.0347838928551667, 0.031262285718929694, 0.02836398007733279, 0.02699196771251519, 0.025628849635548484, 0.02524033136821105]
[-0.037626383473481795, -0.03183971714283873, -0.027379937708911006, -0.02367867061027279, -0.02163876786207286, -0.018819590336618675, -0.017618199273799873]
[-0.2903521482004272, -0.24321610714483333, -0.2087377142810703, -0.1896360199226672, -0.1790080322874848, -0.16737115036445152, -0.16475966863178895]
[0.5063736165265182, 0.40416028285716127, 0.336620062291089, 0.2933213293897272, 0.2723612321379271, 0.25218040966338134, 0.24438180072620014]
10 August 5VA 0.8 PF
(0.16+0.12j)
5 VA 0.8 PF from MSL inductive ~5 VA 0.8 PF
[0.034889839132444145, 0.031326736796737364, 0.028351131528221774, 0.025732251354775304, 0.0240820337462897, 0.020950345303942908, 0.019129713985323482]
[-0.015591868817029342, -0.012921212568458296, -0.010613537820297447, -0.007947037281609504, -0.005653852295755324, -0.0010655619220558965, 0.001150923960077349]
[-1.005110160867556, -0.9126732632032626, -0.8116488684717782, -0.7362677486452247, -0.6979179662537103, -0.6310496546960571, -0.6028702860146765]
[0.5054081311829707, 0.4260787874315417, 0.36738646217970256, 0.3180529627183905, 0.27934614770424465, 0.19993443807794412, 0.17615092396007734]
>Exit code: 0
"""
