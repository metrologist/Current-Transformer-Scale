from __future__ import division
from __future__ import print_function
import math
import GTC as gtc
from twostage import TWOSTAGE
from transformer import TRANSFORMER
from ExcelPython import CALCULATOR

print("Welcome to ctscale")

# construct transformer using dictionaries for windings and cores
ta_primaries = {'P1': [25, 25, 25, 25], 'P2': [5, 5, 5, 5], 'P3': [1, 1, 1, 1, 1]}
ta_secondaries = {'main': 100, 'auxiliary': 100}
ta_cores = {'main': 'could be a polynomial', 'auxiliary': 'could be another polynomial'}
Ta = TRANSFORMER('a', ta_primaries, ta_secondaries, ta_cores, 'current')
print()
print('listing ratios')
print('For Ta')
answer = Ta.dictionary_of_ratios()
print(answer)

tb_primaries = {'P1': [6, 6, 6, 6], 'P2': [3, 3]}
tb_secondaries = {'main': 120, 'auxiliary': 120}
tb_cores = {'main': 'could be a polynomial', 'auxiliary': 'could be another polynomial'}
Tb = TRANSFORMER('b', tb_primaries, tb_secondaries, tb_cores, 'current')

print()
print('listing ratios')

# TODO check for and record agreement/disagreement with old Excel calculations
# The Excel workbooks collect data in a 'Data(xxxx)' sheet where each row has an index.
# This is the likely input data form that python will use.

# labdata = CALCULATOR("Ctcal3_2018.xlsm", "TwoStageResults2018.xlsx")
labdata = CALCULATOR("2018Test.xlsx", "TwoStageResults2018.xlsx")
calrun = TWOSTAGE()

# Work through the buildup process to sort out an efficient set of methods

target_excitation = [125, 120, 100, 60, 40, 20, 10, 5, 1]  # buildup needs a common set of nominal excitation levels
calpage = 'Cal(2018)'  # name of sheet with error measurements
datapage = 'Data(2018)'  # name of sheet with coupling data
# First 5:5 ratio for T_a using the four inner primary windings, primary 1, connected in series
print('Step1')
full_scale = 5.0  # nominal 100 % excitation current

# sign check
block_descriptor = [66, 69, 4, 14]
sign_data1 = labdata.getdata_block(calpage, block_descriptor,)
sign = calrun.signcheck(sign_data1, full_scale, True)
print('the sign is ', sign)

# block_descriptor = [57, 63, 4, 13]  # Ta
# block_descriptor = [155, 162, 4, 14]  # Ta old 2013 data
block_descriptor = [69, 78, 4, 14]  # Ta
copy_data1 = labdata.getdata_block(calpage, block_descriptor)
initial_e1_53, excite1_53, nom_excite1_53 = calrun.ctcompare(copy_data1, full_scale, target_excitation, True)
e1_53 = []
for x in initial_e1_53:
    e1_53.append(x * sign)
print('at 53 Hz')
for q in range(len(e1_53)):
    print(nom_excite1_53[q], repr(e1_53[q].x.real), repr(e1_53[q].x.imag))

block_descriptor = [79, 87, 4, 14]  # Ta
copy_data1 = labdata.getdata_block(calpage, block_descriptor)
initial_e1_47, excite1_47, nom_excite1_47 = calrun.ctcompare(copy_data1, full_scale, target_excitation, True)
e1_47 = []
for x in initial_e1_47:
    e1_47.append(x * sign)
print('at 47 Hz')
for q in range(len(e1_47)):
    print(nom_excite1_47[q], repr(e1_47[q].x.real), repr(e1_47[q].x.imag))


# Second step 20:5 ratio from Ta, primary 1 in parallel, to Tb, primary 1 in series
print('Step2')
# block_descriptor = [112, 118, 4, 13]
#block_descriptor = [169, 177, 4, 14]

# sign check
block_descriptor = [92, 95, 4, 14]
sign_data2 = labdata.getdata_block(calpage, block_descriptor,)
sign = calrun.signcheck(sign_data2, full_scale, True)
print('the sign is ', sign)

block_descriptor = [95, 104, 4, 14]
copy_data2 = labdata.getdata_block(calpage, block_descriptor)
full_scale = 5.0  # nominal 100 % excitation current
initial_e2_53, excite2_53, nom_excite2_53 = calrun.ctcompare(copy_data2, full_scale, target_excitation, True)
e2_53 = []
for x in initial_e2_53:
    e2_53.append(x * sign)
print('at 53 Hz')
for q in range(len(e2_53)):
    print(nom_excite2_53[q], repr(e2_53[q].x.real), repr(e2_53[q].x.imag))


# Third step 100:5 ratio from Tb, primary 1 in parallel, to Ta, primary 2 in parallel
print('Step 3')
# block_descriptor = [164, 170, 4, 13]
# block_descriptor = [184, 191, 4, 14]
# sign check
block_descriptor = [110, 113, 4, 14]
sign_data3 = labdata.getdata_block(calpage, block_descriptor,)
sign = calrun.signcheck(sign_data3, full_scale, True)
print('the sign is ', sign)

block_descriptor = [125, 134, 4, 14]
copy_data3 = labdata.getdata_block(calpage, block_descriptor)
full_scale = 5.0  # nominal 100 % excitation current
initial_e3_53, excite3_53, nom_excite3_53 = calrun.ctcompare(copy_data3, full_scale, target_excitation, True)
e3_53 = []
for x in initial_e3_53:
    e3_53.append(x * sign)
print('at 53 Hz')
for q in range(len(e3_53)):
    print(nom_excite3_53[q], repr(e3_53[q].x.real), repr(e3_53[q].x.imag))


# Fourth step 100:5 from Ta, primary 2 in parallel, to Tb, primary 2 in series
print('Step 4')
# block_descriptor = [244, 250, 4, 13]
# block_descriptor = [248, 255, 4, 14]
# sign check
block_descriptor = [136, 139, 4, 14]
sign_data4 = labdata.getdata_block(calpage, block_descriptor,)
sign = calrun.signcheck(sign_data4, full_scale, True) * -1.0  # because the test was done on Tb!
print('the sign is ', sign)

# note that there are multiple repeats in the raw data that starts at row 139
block_descriptor = [263, 272, 4, 14]
copy_data4 = labdata.getdata_block(calpage, block_descriptor)
full_scale = 5.0  # nominal 100 % excitation current
initial_e4_53, excite4_53, nom_excite4_53 = calrun.ctcompare(copy_data4, full_scale, target_excitation, True)
e4_53 = []
for x in initial_e4_53:
    e4_53.append(x * sign)
print('at 53 Hz')
for q in range(len(e4_53)):
    print(nom_excite4_53[q], repr(e4_53[q].x.real), repr(e4_53[q].x.imag))


# Fifth step a  100:5 outer from Tb, primary 1 in parallel, to Ta ,primary 3 in series
print('Step 5a')
# block_descriptor = [172, 178, 4, 13]
# block_descriptor = [194, 201, 4, 14]
# sign check
block_descriptor = [166, 169, 4, 14]
sign_data5a = labdata.getdata_block(calpage, block_descriptor,)
sign = calrun.signcheck(sign_data5a, full_scale, True)
print('the sign is ', sign)

#again there are multiple repeats so these have been averaged for use lowe in the spreadsheet
block_descriptor = [294, 303, 4, 14]
copy_data5a = labdata.getdata_block(calpage, block_descriptor)
full_scale = 5.0  # nominal 100 % excitation current
initial_e5a_53, excite5a_53, nom_excite5a_53 = calrun.ctcompare(copy_data5a, full_scale, target_excitation, True)
e5a_53 = []
for x in initial_e5a_53:
    e5a_53.append(x * sign)
print('at 53 Hz')
for q in range(len(e5a_53)):
    print(nom_excite5a_53[q], repr(e5a_53[q].x.real), repr(e5a_53[q].x.imag))

# Fifth step b 100:5 from Tb, inner primary 1 in parallel, to Ta ,primary 3 in series
print('Step 5b')
# sign check
block_descriptor = [208, 211, 4, 14]
sign_data5b = labdata.getdata_block(calpage, block_descriptor,)
sign = calrun.signcheck(sign_data5b, full_scale, True)
print('the sign is ', sign)

#again there are multiple repeats so these have been averaged for use lowe in the spreadsheet
block_descriptor = [341, 350, 4, 14]
copy_data5b = labdata.getdata_block(calpage, block_descriptor)
full_scale = 5.0  # nominal 100 % excitation current
initial_e5b_53, excite5b_53, nom_excite5b_53 = calrun.ctcompare(copy_data5b, full_scale, target_excitation, True)
e5b_53 = []
for x in initial_e5b_53:
    e5b_53.append(x * sign)
print('at 53 Hz')
for q in range(len(e5b_53)):
    print(nom_excite5b_53[q], repr(e5b_53[q].x.real), repr(e5b_53[q].x.imag))


# TODO create method that just takes the two blocks and parameters to collapse all the repeated steps below
# get coupling measurements for Ta primary 1
print('Coupling measurements 1a')
# block_descriptor = [84, 87, 3, 13]  # coupling results for inner primary of Ta
block_descriptor = [110, 113, 3, 13]  # coupling results for inner primary of Ta
copy_dataA = labdata.getdata_block(datapage, block_descriptor)
# get volt drops of inner primary Ta
ratio = [4, 25, 100]  #100 / 100  # N2/N1, Ta primary 1
Rs = 0.2  # resistance of secondary burden
rls = 0.268  # leakage resistance of Ta secondary, measured earlier
Vs = 5 * Rs  # nominal voltage across burden at 5 A
# block_descriptor = [51, 55, 4, 8]
block_descriptor = [46, 50, 4, 8]
copy_dataB = labdata.getdata_block(datapage, block_descriptor)
magnetic1a = calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, False)

# get coupling measurements for Tb primary 1
print('Coupling measurements 1b')
# block_descriptor = [139, 143, 3, 13]  # coupling results for inner primary of Tb
block_descriptor = [123, 127, 3, 13]  # coupling results for inner primary of Tb
copy_dataA = labdata.getdata_block(datapage, block_descriptor)
# get volt drops of middle primary Tb
ratio = [5, 6, 120]  # N2/N1, Tb primary 1
Rs = 0.2  # resistance of secondary burden
rls = 0.268  # leakage resistance of Tb secondary, measured earlier
Vs = 5 * Rs  # nominal voltage across burden at 5 A
# block_descriptor = [127, 132, 4, 8]
block_descriptor = [86, 91, 4, 8]
copy_dataB = labdata.getdata_block(datapage, block_descriptor)
magnetic1b = calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, False)

# get coupling measurements for Ta primary 2
print('Coupling measurements 2a')
# block_descriptor = [150, 153, 3, 13]  # coupling results for middle primary of Ta
block_descriptor = [230, 233, 3, 13]  # coupling results for middle primary of Ta
copy_dataA = labdata.getdata_block(datapage, block_descriptor)
# get volt drops of middle primary Ta
ratio = [4, 5, 100]  # N2/N1, Ta primary 2
Rs = 0.2  # resistance of secondary burden
rls = 0.268  # leakage resistance of Ta secondary, measured earlier
Vs = 5 * Rs  # nominal voltage across burden at 5 A
# block_descriptor = [157, 161, 4, 8]
block_descriptor = [220, 224, 4, 8]
copy_dataB = labdata.getdata_block(datapage, block_descriptor)
magnetic2a = calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, False)

# get coupling measurements for Ta primary 3
print('Coupling measurements 3a')
# block_descriptor = [356, 360, 3, 13]  # coupling results for outer primary of Ta
block_descriptor = [238, 242, 3, 13]  # coupling results for outer primary of Ta
copy_dataA = labdata.getdata_block(datapage, block_descriptor)
# get volt drops of outer primary Ta
ratio = [5, 1, 100]  # N2/N1, Ta primary 3
Rs = 0.2  # resistance of secondary burden
rls = 0.268  # leakage resistance of Ta secondary, measured earlier
Vs = 5 * Rs  # nominal voltage across burden at 5 A
# block_descriptor = [190, 195, 4, 8]
block_descriptor = [170, 175, 4, 8]
copy_dataB = labdata.getdata_block(datapage, block_descriptor)
magnetic3a = calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, True)

# get coupling measurements for Tb primary 2
print('Coupling measurements 2b')
# block_descriptor = [226, 227, 3, 13]  # coupling results for outer primary of Tb
block_descriptor = [208, 209, 3, 13]  # coupling results for outer primary of Tb
copy_dataA = labdata.getdata_block(datapage, block_descriptor)
# get volt drops of outer primary Tb
ratio = [2, 3, 120]  # N2/N1, Tb primary 2
Rs = 0.2  # resistance of secondary burden
rls = 0.268  # leakage resistance of Tb secondary, measured earlier
Vs = 5 * Rs  # nominal voltage across burden at 5 A
# block_descriptor = [222, 224, 4, 8]
block_descriptor = [214, 216, 4, 8]
copy_dataB = labdata.getdata_block(datapage, block_descriptor)
# print(len(copy_dataA), len(copy_dataB))
magnetic2b = calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, False)

frequency = 53.0  # Hz, the frequency of calibration
capacitance = 730.8e-12  # farad, capacitance from 100 turn primary to screen
ypg_admit = 1j * 2 * math.pi * frequency * capacitance
k = 4  # number of sections on the primary
z4 = (0.374 + 1j * 0.035)/k  # ohm, primary leakage impedance from Y51, Z51 of 'Calc 50 Hz' in CTCAL2
z3 = 0.2699 + 1j * 0.0021  # ohm, secondary leakage impedance from Z19, AA51 of 'Calc 50 Hz' in CTCAL2
r2 = 0.2  # ohm, secondary burden impedance
ratio = [4, 25, 100]  # 100 / 25  # ratio of series connection
cap_errora = calrun.newcapsp(ypg_admit, z4, r2, z3, ratio)
print('capacitance error = ', cap_errora)
# TODO check result of the capsp method matches Andy's results...No!

frequency = 53.0  # Hz, the frequency of calibration
capacitance = 212.2e-12  # farad, capacitance from 100 turn primary to screen
ypg_admit = 1j * 2 * math.pi * frequency * capacitance
k = 5  # number of sections on the primary
z4 = (0.069818)/k  # ohm, primary leakage impedance from K86 of 'Calc 50 Hz' in CTCAL2 2018
z3 = 0.079 + 1j * 0.00039  # ohm, secondary leakage impedance from S67, X67 of 'Calc 50 Hz' in CTCAL2 2018
r2 = 0.2  # ohm, secondary burden impedance
ratio = [5, 6,100]  # 100 / 25  # ratio of series connection
cap_errorb = calrun.newcapsp(ypg_admit, z4, r2, z3, ratio)
print('capacitance error = ', cap_errorb)

# Use results above to calculate ratios for each transformer
e1 = e1_53
e2 = e2_53
e3 = e3_53
e4 = e4_53
e5a = e5a_53
e5b = e5b_53
t1, t2, t3, t4a, t4b, t5, t6, t7 = calrun.buildup(target_excitation, e1, e2, e3, e4, e5a, e5b, magnetic1a, magnetic2a, magnetic3a,
                                            magnetic1b, magnetic2b, cap_errora, cap_errorb)
# Next we tackle the calibration of Tc against P2a in series.
# Data for this is in a separate spreadsheet, and a different output target that may not be needed
labdata_c = CALCULATOR("1A_CT_cal_2018.xlsm", "TwoStageResults2018_c.xlsx")
calpage_c = '1A CT 2018'
full_scale_c = 1.0  # note that we are now working with 1 A being full scale
# sign check
block_descriptor = [98, 101, 4, 14]  # Tc
sign_data6 = labdata_c.getdata_block(calpage_c, block_descriptor)
sign = calrun.signcheck(sign_data6, full_scale_c, True)
# but in this case the shunt was on the secondary of Tc not Ta and Ta is the reference so it is still the correct sign
# now collect the measured errors

print()
print('This is the Tc analysis')
print('the Tc sign is ',sign)
print()
# note that there are multiple repeats in the raw data that starts at row 101
block_descriptor = [101, 110, 4, 14]
copy_data6 = labdata_c.getdata_block(calpage_c, block_descriptor)
initial_e6_53, excite6_53, nom_excite6_53 = calrun.ctcompare(copy_data6, full_scale_c, target_excitation, True)
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
z4 = (0.374/5.0 + 1j * 0.035)/k  # ohm, primary leakage impedance from Y51, Z51 of 'Calc 50 Hz' in CTCAL2
z3 = 0.2699 + 1j * 0.0021  # ohm, secondary leakage impedance from Z19, AA51 of 'Calc 50 Hz' in CTCAL2
r2 = 0.2  # ohm, secondary burden impedance
ratio = [4, 5, 100]  # 100 / 25  # ratio of series connection
cap_error2a = calrun.newcapsp(ypg_admit, z4, r2, z3, ratio)
print('capacitance error2a = ', cap_error2a)
cap2a = gtc.ucomplex(0+ 0j,(abs(cap_error2a.real), abs(cap_error2a.imag)), label = 'cap2a') # use this estimate as an uncertainty

# Note that there is a factor of 5 difference in the nominal excitation levels of Tc and Ta
# Use simple 2-point extrapolation and interpolation to get the 0.2% to 25% values for P2ap
P2ap_fifth = calrun.one_fifth(t3, target_excitation)  # interpolated P2ap at one fifth excitation
P2as_fifth = []
for x in P2ap_fifth:
    P2as_fifth.append(x - cap2a - magnetic2a) # corrects each point e_series = e_par - cap_sp - mag_sp
# finally assign error to the 5:1 transformer
e6 = []  # this will be the final Tc error
for i in range(len(e6_53)):
    e6.append(P2as_fifth[i] - e6_53[i])  # negative sign on e6 because Ta is in the reference postion


# Now everything about Ta and Tb needs to be manipulated into tables for presenting in a spreadsheet
nom_excite1 = nom_excite1_53
theblock = []
theblock.append(['', 'e1.real', 'e2.real', 'e3.real', 'e4.real', 'e5a.real', 'e5b.real'])
for i in range(len(e1_53)):
    output = []
    output.append(nom_excite1[i])
    output.append(e1[i].x.real)
    output.append(e2[i].x.real)
    output.append(e3[i].x.real)
    output.append(e4[i].x.real)
    output.append(e5a[i].x.real)
    output.append(e5b[i].x.real)
    #output.append('')  # to match width
    theblock.append(output)

theblock.append(['', 'e1.imag', 'e2.imag', 'e3.imag', 'e4.imag', 'e5a.imag', 'e5b.imag'])
for i in range(len(e1)):
    output = []
    output.append(nom_excite1[i])
    output.append(e1[i].x.imag)
    output.append(e2[i].x.imag)
    output.append(e3[i].x.imag)
    output.append(e4[i].x.imag)
    output.append(e5a[i].x.imag)
    output.append(e5b[i].x.imag)
   # output.append('')  # to match width
    theblock.append(output)

theblock.append(['', '1a', '1b', '2a', '3a', '2b', ''])
theblock.append(
    ['real', repr(magnetic1a.x.real), repr(magnetic1b.x.real), repr(magnetic2a.x.real), repr(magnetic3a.x.real),
     repr(magnetic2b.x.real), ''])
theblock.append(
    ['imag', repr(magnetic1a.x.imag), repr(magnetic1b.x.imag), repr(magnetic2a.x.imag), repr(magnetic3a.x.imag),
     repr(magnetic2b.x.imag), ''])
theblock.append(['', 'real', 'imag', '', '', '', ''])
theblock.append(['capa s - p', repr(cap_errora.real), repr(cap_errora.imag), '', '', '', ''])
theblock.append(['capb s - p', repr(cap_errorb.real), repr(cap_errorb.imag), '', '', '', ''])

tableheader = ['Excitation', 'Real', 'Imag', 'U real', 'U imag', 'k real', 'k imag']
theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['P1as', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t1 = labdata.ucomplex_table(target_excitation, t1)
for row in block_t1:
    theblock.append(row)

theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['P1ap', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t2 = labdata.ucomplex_table(target_excitation, t2)
for row in block_t2:
    theblock.append(row)

theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['P2ap', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t3 = labdata.ucomplex_table(target_excitation, t3)
for row in block_t3:
    theblock.append(row)

theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['P3as', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t4a = labdata.ucomplex_table(target_excitation, t4a)
for row in block_t4a:
    theblock.append(row)

theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['P3as', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t4b = labdata.ucomplex_table(target_excitation, t4b)
for row in block_t4b:
    theblock.append(row)

theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['P1bs', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t5 = labdata.ucomplex_table(target_excitation, t5)
for row in block_t5:
    theblock.append(row)

theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['P1bp', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t6 = labdata.ucomplex_table(target_excitation, t6)
for row in block_t6:
    theblock.append(row)

theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['P2bs', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t7 = labdata.ucomplex_table(target_excitation, t7)
for row in block_t7:
    theblock.append(row)

labdata.makeworkbook(theblock, 'my_sheet_name')

print()
print('Example budget, t7[3]')
print()
print(repr(t7[3].real))
for l, u in gtc.rp.budget(t7[3].real):
    print(l, u)

print('For Tb')
answer = Tb.dictionary_of_ratios()
print(answer)




# Now everything about Tc and the Ta reference needs to be presented in a spreadsheet
nom_excite1 = nom_excite6_53
theblock = []
t1 = e6
t2 = P2as_fifth
theblock.append(['', 'e6.real', '-', '-', '-', '-', '-'])
for i in range(len(e1_53)):
    output = []
    output.append(nom_excite1[i])
    output.append(e6_53[i].x.real)
    output.append('')
    output.append('')
    output.append('')
    output.append('')
    output.append('')
    #output.append('')  # to match width
    theblock.append(output)

theblock.append(['', 'e6.imag', '-', '-', '-', '-', '-'])
for i in range(len(e1)):
    output = []
    output.append(nom_excite1[i])
    output.append(e6_53[i].x.imag)
    output.append('')
    output.append('')
    output.append('')
    output.append('')
    output.append('')
   # output.append('')  # to match width
    theblock.append(output)

theblock.append(['', '1a', '1b', '2a', '3a', '2b', ''])
theblock.append(
    ['real', repr(magnetic1a.x.real), repr(magnetic1b.x.real), repr(magnetic2a.x.real), repr(magnetic3a.x.real),
     repr(magnetic2b.x.real), ''])
theblock.append(
    ['imag', repr(magnetic1a.x.imag), repr(magnetic1b.x.imag), repr(magnetic2a.x.imag), repr(magnetic3a.x.imag),
     repr(magnetic2b.x.imag), ''])
theblock.append(['', 'real', 'imag', '', '', '', ''])
theblock.append(['cap2a s - p', repr(cap_error2a.real), repr(cap_error2a.imag), '', '', '', ''])
theblock.append(['capb s - p', repr(cap_errorb.real), repr(cap_errorb.imag), '', '', '', ''])

tableheader = ['Excitation', 'Real', 'Imag', 'U real', 'U imag', 'k real', 'k imag']
theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['Tc', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t1 = labdata.ucomplex_table(target_excitation, t1)
for row in block_t1:
    theblock.append(row)

theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['P2as_fifth', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t2 = labdata.ucomplex_table(target_excitation, t2)
for row in block_t2:
    theblock.append(row)

theblock.append(['', '', '', '', '', '', ''])  # space
theblock.append(['P2ap full', '', '', '', '', '', ''])  # title
theblock.append(tableheader)
block_t3 = labdata.ucomplex_table(target_excitation, t3)
for row in block_t3:
    theblock.append(row)

labdata_c.makeworkbook(theblock, 'my_sheet_name')