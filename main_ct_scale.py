"""
main_ct_scale.py processes all the two-stage calibration data
"""
from __future__ import division
from __future__ import print_function
import math
import GTC as gtc
from ctscale_mod import BUILDUP
from ExcelPython import CALCULATOR

print('Starting ctscale_mod.py')
target_excitation = [125, 120, 100, 60, 40, 20, 10, 5, 1]
build = BUILDUP(['2018Test.xlsx', 'Cal(2018)', 'Data(2018)'], 'TwoStageResults2018_new.xlsx',
                ['1A_CT_cal_2018.xlsm', '1A CT 2018'], 'TwoStageResults2018_c_new.xlsx',
                target_excitation, 5)
print('Step 1')
# First 5:5 ratio for T_a using the four inner primary windings, primary 1, connected in series
e1, excite1, nom_excite = build.step([66, 69, 4, 14], [[69, 78, 4, 14], [79, 87, 4, 14]], True)
e1_53 = e1[0]
e1_47 = e1[1]

print('Step 2')
# 20:5 ratio from Ta, primary 1 in parallel, to Tb, primary 1 in series
e2, excite2, nom_excite = build.step([92, 95, 4, 14], [[95, 104, 4, 14]], True)
e2_53 = e2[0]

print('Step 3')
# 100:5 ratio from Tb, primary 1 in parallel, to Ta, primary 2 in parallel
e3, excite3, nom_excite = build.step([110, 113, 4, 14], [[125, 134, 4, 14]], True)
e3_53 = e3[0]

print('Step 4')
# 100:5 from Ta, primary 2 in parallel, to Tb, primary 2 in series
e4, excite4, nom_excite = build.step([136, 139, 4, 14], [[263, 272, 4, 14]], True)
e4_53 = []
for x in e4[0]:
    e4_53.append(x * -1)  # this is because polarity check was done on Tb rather than Ta

print('Step 5a')
# 100:5 outer from Tb, primary 1 in parallel, to Ta ,primary 3 in series
e5a, excite5a, nom_excite = build.step([166, 169, 4, 14], [[294, 303, 4, 14]], True)
e5a_53 = e5a[0]

print('Step 5b')
# 100:5 from Tb, inner primary 1 in parallel, to Ta ,primary 3 in series
# 100:5 outer from Tb, primary 1 in parallel, to Ta ,primary 3 in series
e5b, excite5b, nom_excite = build.step([208, 211, 4, 14], [[341, 350, 4, 14]], True)
e5b_53 = e5b[0]

# get coupling measurements for Ta primary 1
print('Coupling measurements 1a')
this_ratio = [4, 25, 100]  # 100 / 100  # N2/N1, Ta primary 1
rs = 0.2  # resistance of secondary burden
rrls = 0.268  # leakage resistance of Ta secondary, measured earlier
secI = 5  # nominal 5 A secondary
magnetic1a, magnetic1a_sp = build.magnetic([46, 50, 4, 8], [110, 113, 3, 13], this_ratio, rs, rrls, secI, False)

# get coupling measurements for Tb primary 1
print('Coupling measurements 1b')
this_ratio = [5, 6, 120]  # N2/N1, Tb primary 1
magnetic1b, magnetic1b_sp = build.magnetic([86, 91, 4, 8], [123, 127, 3, 13], this_ratio, rs, rrls, secI, False)

# get coupling measurements for Ta primary 2
print('Coupling measurements 2a')
this_ratio = [4, 5, 100]  # N2/N1, Ta primary 2
magnetic2a, magnetic2a_sp = build.magnetic([220, 224, 4, 8], [230, 233, 3, 13], this_ratio, rs, rrls, secI, False)

# get coupling measurements for Ta primary 3
print('Coupling measurements 3a')
this_ratio = [5, 1, 100]  # N2/N1, Ta primary 3
magnetic3a, magnetic3a_sp = build.magnetic([170, 175, 4, 8], [238, 242, 3, 13], this_ratio, rs, rrls, secI, True)

# get coupling measurements for Tb primary 2
print('Coupling measurements 2b')
this_ratio = [2, 3, 120]  # N2/N1, Tb primary 2
magnetic2b, magnetic2b_sp = build.magnetic([214, 216, 4, 8], [208, 209, 3, 13], this_ratio, rs, rrls, secI, False)

# capacitive error calculation
frequency = 53.0  # Hz, the frequency of calibration
capacitance = 730.8e-12  # farad, capacitance from 100 turn primary to screen
ypg_admit = 1j * 2 * math.pi * frequency * capacitance
k = 4  # number of sections on the primary
z4 = (0.374 + 1j * 0.035) / k  # ohm, primary leakage impedance from Y51, Z51 of 'Calc 50 Hz' in CTCAL2
z3 = 0.2699 + 1j * 0.0021  # ohm, secondary leakage impedance from Z19, AA51 of 'Calc 50 Hz' in CTCAL2
r2 = 0.2  # ohm, secondary burden impedance
ratio = [4, 25, 100]  # 100 / 25  # ratio of series connection
cap_errora, cap_errora_sp = build.calrun.newcapsp(ypg_admit, z4, r2, z3, ratio)
print('capacitance error = ', cap_errora, cap_errora_sp)

frequency = 53.0  # Hz, the frequency of calibration
capacitance = 212.2e-12  # farad, capacitance from 100 turn primary to screen
ypg_admit = 1j * 2 * math.pi * frequency * capacitance
k = 5  # number of sections on the primary
z4 = (0.069818) / k  # ohm, primary leakage impedance from K86 of 'Calc 50 Hz' in CTCAL2 2018
z3 = 0.079 + 1j * 0.00039  # ohm, secondary leakage impedance from S67, X67 of 'Calc 50 Hz' in CTCAL2 2018
r2 = 0.2  # ohm, secondary burden impedance
ratio = [5, 6, 100]  # 100 / 25  # ratio of series connection
cap_errorb, cap_errorb_sp = build.calrun.newcapsp(ypg_admit, z4, r2, z3, ratio)
print('capacitance error = ', cap_errorb, cap_errorb_sp)

# Use results above to calculate ratios for each transformer
e1 = e1_53
e2 = e2_53
e3 = e3_53
e4 = e4_53
e5a = e5a_53
e5b = e5b_53
t1, t2, t3, t4a, t4b, t5, t6, t7 = build.calrun.buildup(target_excitation, e1, e2, e3, e4, e5a, e5b, magnetic1a,
                                                        magnetic2a, magnetic3a,
                                                        magnetic1b, magnetic2b, cap_errora, cap_errorb)

# Next we tackle the calibration of Tc against P2a in series.
# Data for this is in a separate spreadsheet, and a different output target that may not be needed

full_scale_c = 1.0  # note that we are now working with 1 A being full scale
# sign check
block_descriptor = [98, 101, 4, 14]  # Tc
sign_data6 = build.labdata_c.getdata_block(build.calpage_c, block_descriptor)
sign = build.calrun.signcheck(sign_data6, full_scale_c, True)
# but in this case the shunt was on the secondary of Tc not Ta and Ta is the reference so it is still the correct sign
# now collect the measured errors

print()
print('This is the Tc analysis')
print('the Tc sign is ', sign)
print()
# note that there are multiple repeats in the raw data that starts at row 101
block_descriptor = [101, 110, 4, 14]
copy_data6 = build.labdata_c.getdata_block(build.calpage_c, block_descriptor)
initial_e6_53, excite6_53, nom_excite6_53 = build.calrun.ctcompare(copy_data6, full_scale_c, target_excitation, True)
e6_53 = []
for x in initial_e6_53:
    e6_53.append(x * sign)
print('at 53 Hz')
for q in range(len(e6_53)):
    print(nom_excite6_53[q], repr(e6_53[q].x.real), repr(e6_53[q].x.imag))

# Now need the series error for P2as, based on P2ap (magnetic2a and cap_error_2a)
print('magnetic2a ', magnetic2a)
frequency = 53.0  # Hz, the frequency of calibration
capacitance = 197.39e-12  # farad, capacitance from 20 turn primary to screen
ypg_admit = 1j * 2 * math.pi * frequency * capacitance
k = 4  # number of sections on the primary
# z4 has not been measured and is instead estimated as being P1 resistance value divided by 5, but a retain high full leakage impedance
z4 = (0.374 / 5.0 + 1j * 0.035) / k  # ohm, primary leakage impedance from Y51, Z51 of 'Calc 50 Hz' in CTCAL2
z3 = 0.2699 + 1j * 0.0021  # ohm, secondary leakage impedance from Z19, AA51 of 'Calc 50 Hz' in CTCAL2
r2 = 0.2  # ohm, secondary burden impedance
ratio = [4, 5, 100]  # 100 / 25  # ratio of series connection
cap_error2a, cap_error2a_sp = build.calrun.newcapsp(ypg_admit, z4, r2, z3, ratio)
print('capacitance error2a = ', cap_error2a, cap_error2a_sp)
cap2a = gtc.ucomplex(0 + 0j, (abs(cap_error2a.real), abs(cap_error2a.imag)),
                     label='cap2a')  # use this estimate as an uncertainty

# Note that there is a factor of 5 difference in the nominal excitation levels of Tc and Ta
# Use simple 2-point extrapolation and interpolation to get the 0.2% to 25% values for P2ap
P2ap_fifth = build.calrun.one_fifth(t3, target_excitation)  # interpolated P2ap at one fifth excitation
P2as_fifth = []
for x in P2ap_fifth:
    P2as_fifth.append(x - cap2a - magnetic2a)  # corrects each point e_series = e_par - cap_sp - mag_sp
# finally assign error to the 5:1 transformer
e6 = []  # this will be the final Tc error
for i in range(len(e6_53)):
    e6.append(P2as_fifth[i] - e6_53[i])  # negative sign on e6 because Ta is in the reference postion

# Now return to Ta and Tb
# Additional ratios can be calculated from t1...t7 for the connections that were not part of the buildup.
# Where values of capacitive or magnetic coupling errors have not been measured an uncertain estimate is used.
# Not measured is 'cap2b' (historically not done ... do next time)
cap_error2b = gtc.ucomplex(complex(0,0), (1e-7, 1e-7), label = 'cap_error2b')  # 0.1 ppm in each axis

t8, t9, t10, t11, t12 = build.calrun.extra_ratios(t2, t3, t4b, t7, magnetic1a, magnetic1a_sp, magnetic2a, magnetic2a_sp,
                                                  magnetic3a, magnetic2b, cap_errora, cap_errora_sp, cap_error2a_sp,
                                                  cap_error2b)

print('Create tables for Ta and Tb in the output Excel files')
errors = [e1, e2, e3, e4, e5a, e5b]
err_mag = [magnetic1a, magnetic1b, magnetic2a, magnetic3a, magnetic2b]
err_cap = [cap_errora, cap_errorb]
err_t = [t1, t2, t3, t4a, t4b, t5, t6, t7, t8, t9, t10, t11, t12]
build.report_ab(target_excitation, errors, err_mag, err_cap, err_t)

print('Create tables for Tc measured against Ta')
errors = [e1_53, e6_53]
err_t = [e6, P2as_fifth, t3]
err_cap = [cap_error2a, cap_errorb]
build.report_ac(target_excitation, errors, err_mag, err_cap, err_t)

# Finally example uncertainty budgets can be generated and put in Excel for review
# Reports are built using the 't*' values
# Force labels on the derived uncertain numbers
index = 8
t1[index].label = ('t1_' + repr(index))
t2[index].label = ('t2_' + repr(index))
t3[index].label = ('t3_' + repr(index))
t4a[index].label = ('t4a_' + repr(index))
t4b[index].label = ('t4b_' + repr(index))
t5[index].label = ('t5_' + repr(index))
t6[index].label = ('t6_' + repr(index))
t7[index].label = ('t7_' + repr(index))
P2as_fifth[index].label = ('P2as_fifth_' + repr(index))
e6[index].label = ('e6_' + repr(index))
uncertain_sheet = CALCULATOR('Dummy.xlsx', 'uncertainties.xlsx')  # The need for Dummy.xlsx is an artefact of the class

budgets = uncertain_sheet.budget_table([t1[index].real,t2[index].real, t3[index].real ,t4a[index].real, t4b[index].real,
                                        t5[index].real, t6[index].real, t7[index].real, P2as_fifth[index].real,
                                        e6[index].real,
                                        t1[index].imag,t2[index].imag, t3[index].imag ,t4a[index].imag, t4b[index].imag,
                                        t5[index].imag, t6[index].imag, t7[index].imag,P2as_fifth[index].imag,
                                        e6[index].imag], 0.1)
# print(budgets)
print('\n', 'Check available series/parallel corrections')
print('cap_error2a_sp', cap_error2a_sp)
print('cap_errora_sp', cap_errora_sp)
print('cap_errorb_sp', cap_errorb_sp)
print('magnetic1a_sp', magnetic1a_sp)
print('magnetic1b_sp', magnetic1b_sp)
print('magnetic2a_sp', magnetic2a_sp)
print('magnetic2b_sp', magnetic2b_sp)
print('magnetic3a_sp', magnetic3a_sp)
