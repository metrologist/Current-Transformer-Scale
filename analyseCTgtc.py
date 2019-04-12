from __future__ import division
"""
Use least squares to choose best value of secondary leakage impedance
The CT class has been modified here to take in the real and imaginary part
of the secondary leakage separately, rather than as a complex number.
"""
import modelCT
from scipy.optimize import minimize

import math
import GTC
import ExcelPython


if __name__ == "__main__":
    #use fitted impedance coefficients to define behviour of core impedance
    this_core = [11.7204144319303, -0.0273771774408702, 29.7364068940134, 3.90272331971033, 0.000243735060998679, 5.8544055409592]
    this_CT = modelCT.CT(this_core)
    	
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
    hessian = result.hess_inv #inverse Hessian is estimate of variance/covariance matrix
    sum_square = result.fun#final rss
    #scaling the inverse hessian by multiplying by the residuals
    sa0 = math.sqrt(hessian[0,0]*sum_square)
    sa1 = math.sqrt(hessian[1,1]*sum_square)
    a = result.x
##     zs = GTC.ucomplex(a[0]+1j*a[1],(hessian[0,0]*sum_square,hessian[0,1]*sum_square,hessian[1,0]*sum_square,hessian[1,1]*sum_square),138)
    zs = GTC.ucomplex(a[0]+1j*a[1],(hessian[0,0]*sum_square,hessian[1,1]*sum_square),df=138,label='zs')

    #use fitted impedance coefficients to define behviour of core impedance
    #should look at GTC output when model inputs have uncertainty from the fits
    a1 = GTC.ureal(11.7204144319303,0.346683597109428,df=34,label ='a1',dependent = True)
    a2 = GTC.ureal(-0.0273771774408702,0.0136282110753501,df=34,label ='a2',dependent = True)
    a3 = GTC.ureal(29.7364068940134,0.695696459888704,df=34,label ='a3',dependent = True)
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
    this_CT_gtc = modelCT.CT(this_core_gtc)
    x = GTC.ureal(5,0.05,df=5,label='excitation') #excitation set with a precision of 1%
    core_imp = this_CT_gtc.coreZgtc(x, 50)
##     bdn = GTC.ucomplex(0.16 +1j*0.12,(0.01,0.01),df=5, label='burden')
    power_factor = GTC.ureal(0.8,0.1,df=5, label ='PF')
    volt_amp = GTC.ureal(0.5,0.1,df = 5,label = 'VA')
    bdn = this_CT_gtc.burdenVAgtc(volt_amp,power_factor,5.0)
    final = this_CT_gtc.error_z(core_imp,zs,bdn)*100
    print 'final',GTC.summary(final)
##     print this_CT_gtc.burdenVA(5,0.8,5)
    for l,u in GTC.reporting.budget(final, reverse = True,trim = 0):
        print "%s: %G" % (l,u)
        
    print GTC.component(final.imag,zs.imag)
    print GTC.component(final.real,volt_amp.real)
    print GTC.reporting.sensitivity(final,volt_amp.real)

    
##     print my_copy_data
    #put the block, unchanged into the output spreadsheet
##     ianz.makeworkbook(my_copy_data, 'fromS21982_1_turn')
"""
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

## excite = 60
## core = this_CT.coreZ(excite,50)
## print
## print 'calculate error with zH'
## complex_error = this_CT.error_z(core, zs, zH)*100
## print repr(complex_error)
## print complex_error.u
## print
## print 'calculate error with zB'
## complex_error2 = this_CT.error_z(core, zs, zB)*100
## print repr(complex_error2)
## print complex_error2.u
## print
## print'calculate difference'
## correction = complex_error - complex_error2
## print repr(correction)
## print correction.u
print
print 'Now calculate corrections to quote results at exact nominal burden'


#manual selection of which blocks will be corrected to a nominal burden.
#using 'complex(zs)' breaks the GTC chain, which keeps the number format simple
#the uncertainty in the correction is negligible as same coefficients used

target_burden2 = this_CT.burdenVA(5.0,1.0,5.0)
table2 = this_CT.burden_correction_z(target_burden2,complex(zs),zB,x,a2,b2)

target_burden5 = this_CT.burdenVA(1.25,1.0,5.0)
table5 = this_CT.burden_correction_z(target_burden5,complex(zs),zE,x,a5,b5)

target_burden7 = this_CT.burdenVA(5.0,0.8,5.0)
table7 = this_CT.burden_correction_z(target_burden7,complex(zs),zH,x,a7,b7)

target_burden8 = this_CT.burdenVA(5.0,1.0,5.0)
table8 = this_CT.burden_correction_z(target_burden8,complex(zs),zB,x,a8,b8)

target_burden9 = this_CT.burdenVA(5.0,0.8,5.0)
table9 = this_CT.burden_correction_z(target_burden9,complex(zs),zH,x,a9,b9)



#Output these results to a spreadsheet.
#Can either add to my_copy_data, creating a single block structure for outputing
#or could address the spreadsheet cell by cell ... all options are cumbersome.
#the corrected results are two additional columns, so probably would like a 
#column writing option
print my_copy_data

for i in range(len(my_copy_data)):
    for j in range(3):
        my_copy_data[i].append(None)# appends 3 blank columns

r1 = 125
r2 = 131
for i in range(r1-1,r2): #-1 makes it match the right row number:)
    print i
    end = len(my_copy_data[i])
    my_copy_data[i][end-2] = table2[0][i-r1+1]#the +1 makes it work :)
    my_copy_data[i][end-1] = table2[1][i-r1+1]


r1 = 91
r2 = 97
for i in range(r1-1,r2): #-1 makes it match the right row number:)
    end = len(my_copy_data[i])
    my_copy_data[i][end-2] = table5[0][i-r1+1]
    my_copy_data[i][end-1] = table5[1][i-r1+1]

r1 = 159
r2 = 165
for i in range(r1-1,r2): #-1 makes it match the right row number:)
    end = len(my_copy_data[i])
    my_copy_data[i][end-2] = table7[0][i-r1+1]
    my_copy_data[i][end-1] = table7[1][i-r1+1]

r1 = 210
r2 = 216
for i in range(r1-1,r2): #-1 makes it match the right row number:)
    end = len(my_copy_data[i])
    my_copy_data[i][end-2] = table8[0][i-r1+1]
    my_copy_data[i][end-1] = table8[1][i-r1+1]

r1 = 227
r2 = 233
for i in range(r1-1,r2): #-1 makes it match the right row number:)
    end = len(my_copy_data[i])
    my_copy_data[i][end-2] = table9[0][i-r1+1]
    my_copy_data[i][end-1] = table9[1][i-r1+1]
    


print my_copy_data
ianz.makeworkbook(my_copy_data,'my_sheet_name')
"""
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

