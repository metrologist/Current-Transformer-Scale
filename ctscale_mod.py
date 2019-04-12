from __future__ import division
from __future__ import print_function
import math
import GTC as gtc
from twostage import TWOSTAGE
from ExcelPython import CALCULATOR

class BUILDUP(object):
    """
    uses selected data from Excel spreadhseets to calibrate Ta, Tb and Tc.
    """
    def __init__(self, spreadin1, spreadout1, spreadin2, spreadout2, excitation_levels, sec100):
        """

        :param spreadin1: input spreadsheet list for Ta and Tb [file name, error measure sheet, coupling sheet]
        :param spreadout1: output spreadsheet for Ta and Tb
        :param spreadin2: input spreadsheet list for Tc [file name, error measure sheet]
        :param spreadout2: output spreadsheet for Tc
        :param excitation_levels: list of % excitation levels that must be common throughout
        :param sec100: the secondary current at 100% excitation
        """
        self.labdata = CALCULATOR(spreadin1[0], spreadout1)
        self.calpage = spreadin1[1]
        self.datapage = spreadin1[2]
        self.calrun = TWOSTAGE()
        self.labdata_c = CALCULATOR(spreadin2[0], spreadout2)
        self.calpage_c = spreadin2[1]
        self.target_excitation = excitation_levels
        self.full_scale = sec100

    def step(self, sign_block, error_block, bool):
        """
        The location of data blocks in the spreadsheet can vary depending on various investigations that might
        result during a calibration run. Once the data for a particular step is identified, this method does
        the necessary calculations.

        :param sign_block: location of sign check data [row start, 1+row finish, column start column finish]
        :param error_block: list of blocks with measurement data (typically 53 Hz and 47 Hz data)
        :param bool: boolean for correct columns
        :return: a list of lists of calculated errors, a matching list of lists of actual excitations, a single list
        of nominal excitation.
        """
        # sign check
        block_descriptor = sign_block
        sign_data1 = self.labdata.getdata_block(self.calpage, block_descriptor, )
        sign = self.calrun.signcheck(sign_data1, self.full_scale, True)
        e = []  # will be a list of the lists of errors
        exc = []

        for i in range(len(error_block)):
            block_descriptor = error_block[i]
            copy_data1 = self.labdata.getdata_block(self.calpage, block_descriptor)
            initial_e, excite, nom_excite = self.calrun.ctcompare(copy_data1, self.full_scale, self.target_excitation, bool)
            e_freq = []  # this is the errors at the test frequency modified to have the correct sign
            for x in initial_e:
                e_freq.append(x * sign)
            e.append(e_freq)
            exc.append(excite)
        return e, exc, nom_excite

    def magnetic(self, block_share, block_couple, ratio, Rs, rls, Is, bool):
        """

        :param block_share: row-column list for measured volt drops on the primary sections
        :param block_couple: row-column list for coupling measurements
        :param ratio: is a list [k, m, n] of k primary sections of m turns with a secondary of n turns
        :param Rs: resistance of the burden on the secondary
        :param rls: secondary leakage resistance
        :param Is: 100 % value of secondary current
        :param bool: boolean when set to True the vdrops are measured across each individual winding
        :return: ucomplex value of series - parallel error
        """
        copy_dataA = self.labdata.getdata_block(self.datapage, block_couple)
        # get volt drops of primary
        Vs = Is * Rs  # nominal voltage across burden at 5 A
        copy_dataB = self.labdata.getdata_block(self.datapage, block_share)
        magnetic1a = self.calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, bool)
        return magnetic1a

    def build_block(self, x_excite, ablock, tbl_head, tfm_label, tfm):
        """
        For each transformer configuration this builds a results table and appends it to 'ablock', which is the
        block that is ultimately passed to the output spreadsheet

        :param x_excite: list of excitation levels
        :param ablock: a block of equal length rows for sending to the output spreadsheet
        :param tbl_head: heading for the columns of the tables
        :param tfm_label: the name of the transformer ratio
        :param tfm: the eorror list for the transformer
        :return: ablock, which can be further appended by repeat calls to this method
        """
        ablock.append(['', '', '', '', '', '', ''])  # space
        ablock.append([tfm_label, '', '', '', '', '', ''])  # title
        ablock.append(tbl_head)
        tfm_block = self.labdata.ucomplex_table(x_excite, tfm)
        for row in tfm_block:
            ablock.append(row)

        return ablock

    def report_ab(self, nom_excite1, errors, err_mag, err_cap, err_t):
        """
        Gathers all errors for Ta and Tb for presentation in a spreadsheet for creating a final calibration report.

        :param nom_excite1: list of the nominal % excitation levels
        :param errors: a list of the measured errors [e1, e2, e3, e4, e5a, e5b] in the buildup
        :param err_mag: a list of the magnetic series-parallel errors [magnetic1a,magnetic1b,magnetic2a,magnetic3a,magnetic2b]
        :param err_cap: a list of the capacitive series-parallel errors [cap_errora, cap_errorb]
        :param err_t: a list of the calculated errors [t1, t2, t3, t4a, t4b, t5, t6, t7]
        :return: report blocks for Ta and Tb are placed in spreadout2
        """

        target_excitation = nom_excite1
        e1, e2, e3, e4, e5a, e5b = errors[0], errors[1], errors[2], errors[3], errors[4], errors[5]
        t1, t2, t3, t4a, t4b, t5, t6, t7 = err_t[0],err_t[1],err_t[2],err_t[3],err_t[4],err_t[5],err_t[6], err_t[7]
        t8, t9, t10, t11, t12 = err_t[8], err_t[9], err_t[10], err_t[11], err_t[12]  # the extra ratios
        cap_errora, cap_errorb = err_cap[0], err_cap[1]
        magnetic1a,magnetic1b,magnetic2a,magnetic3a,magnetic2b = err_mag[0],err_mag[1],err_mag[2],err_mag[3],err_mag[4]
        theblock = []
        theblock.append(['', 'e1.real', 'e2.real', 'e3.real', 'e4.real', 'e5a.real', 'e5b.real'])
        for i in range(len(e1)):
            output = []
            output.append(nom_excite1[i])
            output.append(e1[i].x.real)
            output.append(e2[i].x.real)
            output.append(e3[i].x.real)
            output.append(e4[i].x.real)
            output.append(e5a[i].x.real)
            output.append(e5b[i].x.real)
            # output.append('')  # to match width
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
        self.build_block(nom_excite1,theblock, tableheader, 'P1as',t1)
        self.build_block(nom_excite1, theblock, tableheader, 'P1ap', t2)
        self.build_block(nom_excite1, theblock, tableheader, 'P2ap', t3)
        self.build_block(nom_excite1, theblock, tableheader, 'P3as', t4a)
        self.build_block(nom_excite1, theblock, tableheader, 'P3as', t4b)
        self.build_block(nom_excite1, theblock, tableheader, 'P1bs', t5)
        self.build_block(nom_excite1, theblock, tableheader, 'P1bp', t6)
        self.build_block(nom_excite1, theblock, tableheader, 'P2bs', t7)

        self.build_block(nom_excite1, theblock, tableheader, 'P3ap', t8)
        self.build_block(nom_excite1, theblock, tableheader, 'P2as', t9)
        self.build_block(nom_excite1, theblock, tableheader, 'P2asp', t10)
        self.build_block(nom_excite1, theblock, tableheader, 'P1asp', t11)
        self.build_block(nom_excite1, theblock, tableheader, 'P2bp', t12)

        self.labdata.makeworkbook(theblock, 'my_sheet_name')

    def report_ac(self, nom_excite1, errors, err_mag, err_cap, err_t):
        """
        Gathers all errors for Tc for presentation in a spreadsheet for creating a final calibration report.

        :param nom_excite1: list of the nominal % excitation levels
        :param errors: list of measured errors [e1_53, e6_53]
        :param err_mag: list of magnetic series-parallel errors [magnetic1a,magnetic1b,magnetic2a,magnetic3a,magnetic2b]
        :param err_cap: list of capacitive series-parallel [cap_error2a, cap_errorb]
        :param err_t: list of calculated errors [t1, t2, t3]
        :return: a Tc report is placed in output spreadsheet spreadout2
        """
        target_excitation = nom_excite1
        t1, t2, t3 = err_t[0], err_t[1], err_t[2]
        e1_53, e6_53 = errors[0], errors[1]
        cap_error2a, cap_errorb = err_cap[0], err_cap[1]
        magnetic1a,magnetic1b,magnetic2a,magnetic3a,magnetic2b = err_mag[0],err_mag[1],err_mag[2],err_mag[3],err_mag[4]
        theblock = []
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
            # output.append('')  # to match width
            theblock.append(output)

        theblock.append(['', 'e6.imag', '-', '-', '-', '-', '-'])
        for i in range(len(e1_53)):
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
        self.build_block(nom_excite1, theblock, tableheader, 'Tc', t1)
        self.build_block(nom_excite1, theblock, tableheader, 'P2as_fifth', t2)
        self.build_block(nom_excite1, theblock, tableheader, 'P2ap full', t3)

        self.labdata_c.makeworkbook(theblock, 'my_sheet_name')


if __name__ == "__main__":
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
    print('capacitance error = ', cap_errora)

    frequency = 53.0  # Hz, the frequency of calibration
    capacitance = 212.2e-12  # farad, capacitance from 100 turn primary to screen
    ypg_admit = 1j * 2 * math.pi * frequency * capacitance
    k = 5  # number of sections on the primary
    z4 = (0.069818) / k  # ohm, primary leakage impedance from K86 of 'Calc 50 Hz' in CTCAL2 2018
    z3 = 0.079 + 1j * 0.00039  # ohm, secondary leakage impedance from S67, X67 of 'Calc 50 Hz' in CTCAL2 2018
    r2 = 0.2  # ohm, secondary burden impedance
    ratio = [5, 6, 100]  # 100 / 25  # ratio of series connection
    cap_errorb, caperrorb_sp = build.calrun.newcapsp(ypg_admit, z4, r2, z3, ratio)
    print('capacitance error = ', cap_errorb)

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
    print('capacitance error2a = ', cap_error2a)
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

    print('Create tables for Ta and Tb in the output Excel files')
    errors = [e1, e2, e3, e4, e5a, e5b]
    err_mag = [magnetic1a,magnetic1b,magnetic2a,magnetic3a,magnetic2b]
    err_cap = [cap_errora, cap_errorb]
    err_t = [t1, t2, t3, t4a, t4b, t5, t6, t7]
    build.report_ab(target_excitation, errors, err_mag, err_cap, err_t)

    print('Create tables for Tc measured against Ta')
    errors = [e1_53, e6_53]
    err_t = [e6, P2as_fifth, t3]
    err_cap = [cap_error2a, cap_errorb]
    build.report_ac(target_excitation, errors, err_mag, err_cap, err_t)

