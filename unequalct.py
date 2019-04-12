from __future__ import division
from __future__ import print_function
import math
import GTC as gtc
from ExcelPython import CALCULATOR

"""
unequalct.py deals with data from the unequal ratios method in technical procedure E.059.004.
Initially this is just a common sense check on the spreadsheet, but should be developed into the default calculation
method. 
"""
labdata = CALCULATOR("unequal_test.xlsx", "UnequalResults2018.xlsx")
calpage = 'Data'
block_descriptor = [13, 14, 3, 22]
data1 = labdata.getdata_block(calpage, block_descriptor)
data = data1[0]  # 'returns a list of rows, in this case just one row
# for clarity now manually assign names tho the block that has been read in.

shuntv = data[0]
vx = data[1]
vy = data[2]
sdx = data[3]
sdy = data[4]
sdshunt = data[5]
temprt = data[6]
ns = data[9]
nt =data[10]
alpha = data[11]
rsnom = data[12]
rtnom = data[13]
sign = data[14]


print ('shuntv =', shuntv)
print('vx =', vx)
print('vy =', vy)
print('sdx =', sdx)
print('sdy =', sdy)
print('sdshunt =', sdshunt)
print('temprt =', temprt)
print('ns =', ns)
print('nt =', nt)
print('alpha =', alpha)
print('rsnom =', rsnom)
print('rtnom =', rtnom)
print('sign =', sign)


# now ready to calculate
Ns = ns  # for now assume no error in the reference
dV = sign * (vx + 1j*vy)
rt_rs = (alpha / ns) * (1 - (dV / (alpha * shuntv)))
print('rt_rs =', rt_rs)

print()
print()


# calculate ratio of CT under test
block_descriptor = [37, 38, 3, 22]
data1 = labdata.getdata_block(calpage, block_descriptor)
data = data1[0]  # 'returns a list of rows, in this case just one row
# for clarity now manually assign names tho the block that has been read in.

shuntv = data[0]
vx = data[1]
vy = data[2]
sdx = data[3]
sdy = data[4]
sdshunt = data[5]
temprt = data[6]
ns = data[9]
nt =data[10]
alpha = data[11]
rsnom = data[12]
rtnom = data[13]
sign = data[14]

print ('shuntv =', shuntv)
print('vx =', vx)
print('vy =', vy)
print('sdx =', sdx)
print('sdy =', sdy)
print('sdshunt =', sdshunt)
print('temprt =', temprt)
print('ns =', ns)
print('nt =', nt)
print('alpha =', alpha)
print('rsnom =', rsnom)
print('rtnom =', rtnom)
print('sign =', sign)


rt_rs = 0.5 * (1.000707 +1j * 0.000119)  # taken from spreadsheet
print('rt_rs =', rt_rs)
sign = -1*sign  # based on polarity test
# now ready to calculate
Ns = ns  # for now assume no error in the reference
dV = sign * (vx + 1j*vy)
Nt = Ns * ((1 / alpha) * rt_rs * (1 / (1 - (dV / (alpha * shuntv)))))
print('Nt =', Nt)
et =(nt/Nt - 1)
print('error =', et)
print(et.real * 1000)
print(et.imag * 1000)