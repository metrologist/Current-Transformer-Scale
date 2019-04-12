from __future__ import division
"""
Calculating value of burden by perturbing secondary current with a resistor
across output terminals of a CT
This version is a bit more time efficient with accessing the worksheet.
"""
import ExcelPython
import modelCT  # just for burden VA calculations
import arnold  # uncertainty terms for dials
import GTC as gtc
import math

class BURDEN(object):
    """
    Simple measurement method based on perturbing the burden by connecting an impedance in parallel.
    Changes in the secondary current error of the CT are then used to calculate the burden.
    """
    def __init__(self, excelfile, datasheet, start_stop, resistors):
        """

        :param excelfile: contains the worksheet 'datasheet' that contains all the essential data
        :param datasheet: the sheet name
        :param start_stop: [start row, stop row, start column, stop colunm]
        :param resistors: dictionary of values of the perturbing resistors
        """
        self.excelfile = excelfile
        self.datasheet = datasheet
        self.data_store = ExcelPython.CALCULATOR(self.excelfile, "myBurdenResults.xlsx")
        self.perturbr = resistors  # as a dictionary
        self.arnie = arnold.ARNOLD('source', 'output)')  # dummy source and output as not needed
        self.start_stop = start_stop  # [start row, stop row, start col, stop col]
        self.datalist = self.data_store.getdata_block(self.datasheet, start_stop)
        self.maxfreq = 1.5e-2# 1.5% max allowed deviation

    def burdenZ(self, e1, e2, referenceR):
        """
        :param e1: complex error with referenceR
        :param e2: complex error without referenceR
        :param referenceR: the resistor put across the burden
        :return: burden impedance
        """
        e_delta = (e1 - e2) / 100
        return e_delta / (1 - e_delta) * referenceR

    def tuple_to_complex(self, a):
        """
        :param a: tuple x, y
        :return: x + jy
        """
        return a[0] + 1j * a[1]

    def getblock(self, rowcol):
        """
        extracts content from the datablock lists that were read in from the wosksheet
        :param rowcol: list of [start row, stop row, startcol, stopcol]
        :return: a list of rows as lists
        """
        outputblock = []
        for i in range(rowcol[0], rowcol[1]):
            this_row = []
            for j in range(rowcol[2], rowcol[3]):
                this_row.append(self.datalist[i-1][j-1])  # the '-1' is an artefact of the indicies being excel r1, c1
            outputblock.append(this_row)
        return outputblock

    def get_test_result(self, a):
        """
        :param a: first row of test
        :return: list of (ultimately ureals) [complex error 1, complex error 2, resistance]
        """
        # first get relevant resistor
        resistor = self.perturbr[str(self.getblock([a +1, a + 3 , 12, 13])[0][0])]
        multiplier = float(self.getblock([a , a + 1 , 7, 8])[0][0])

        errors = self.getblock([a , a + 2 , 10, 12])
        error1 = self.tuple_to_complex(errors[0])
        u_resolutionR = gtc.ureal(0,self.arnie.resolutionR(error1.real, multiplier), 12, 'resR' + repr(error1.real))
        u_correctionR = gtc.ureal(0, self.arnie.corrnR(error1.real, multiplier), 12, 'corrnR' + repr(error1.real))
        u_resolutionM = gtc.ureal(0,self.arnie.resolutionM(error1.imag, multiplier), 12, 'resM' + repr(error1.imag))
        u_correctionM = gtc.ureal(0,self.arnie.corrnM(error1.imag, multiplier), 12, 'corrnM' + repr(error1.imag))

        error = error1.real + u_resolutionR + u_correctionR + 1j*(error1.imag + u_resolutionM + u_correctionM )
        error1 = error

        error2 = self.tuple_to_complex(errors[1])
        u_resolutionR = gtc.ureal(0, self.arnie.resolutionR(error2.real, multiplier), 12, 'resR' + repr(error2.real))
        u_correctionR = gtc.ureal(0, self.arnie.corrnR(error2.real, multiplier), 12, 'corrnR' + repr(error2.real))
        u_resolutionM = gtc.ureal(0, self.arnie.resolutionM(error2.imag, multiplier), 12, 'resM' + repr(error2.imag))
        u_correctionM = gtc.ureal(0, self.arnie.corrnM(error2.imag, multiplier), 12, 'corrnM' + repr(error2.imag))

        error = error2.real + u_resolutionR + u_correctionR + 1j * (error2.imag + u_resolutionM + u_correctionM)
        error2 = error

        the_burden = self.burdenZ(error1, error2, resistor)
        frequency_factor = gtc.ureal(0, self.maxfreq / 3.0, 50, label='frequency')
        the_burden = the_burden + 1j * the_burden.imag * frequency_factor  # the addition is nominally zero

        return the_burden

    def unique_list(self, input):
        """

        :param input: a list with duplicate entries
        :return: a list with no duplicates
        """
        output = [input[0]]  #start with the first entry loaded
        for x in input:
            if x not in output:
                output.append(x)  # only add if it isn't already in the list
        return output

if __name__ == "__main__":
    #use fitted impedance coefficients to define behviour of core impedance
    this_core = [11.7204144319303, -0.0273771774408702, 29.7364068940134, 3.90272331971033, 0.000243735060998679, 5.8544055409592]
    this_CT = modelCT.CT(this_core)  # a convenience CT

    w = 2*math.pi*50  # angular mains frequency

    # perturbing resistors values from S22008_burden_support.xlsx
    ohm10 = gtc.ucomplex((10.0112386 + 1j * w * 1724.3e-9 ), (0.0005995/2.0, 52e-9/2.0 ), 50, label = '10 ohm' )
    ohm50 = gtc.ucomplex((50.0046351 + 1j * w * 2590.5e-9 ), (0.0050221/2.0, 426.2e-9/2.0 ), 50, label = '50 ohm' )
    ohm100 = gtc.ucomplex((100.0048255 + 1j * w * -4140.8e-9 ), (0.0057092/2.0, 520.1e-9/2.0 ), 50, label = '100 ohm' )
    ohm1000 = gtc.ureal(1000, 0.01, 12, label = '1000 ohm') #not measured on UB
    check_r = {'10': ohm10, '50': 50.0, '100': 100.0, '1000': 1000.0}  # dictionary of values of perturbing resistors
    # analyse data from Excel sheet
    myburden = BURDEN('S22008_py_Burden_5AmpRatio Template V5.0_draft.xlsm', 'ForPython', [1, 128, 1, 17], check_r)
    mydata = myburden.datalist
    print('starting run')
    test_sets = [4, 6, 11, 13, 18, 25, 27, 32, 34, 39, 41, 46, 48, 53, 55, 60, 62, 67, 69,74, 76, 81, 83, 88,
                 90, 95, 97, 102, 104, 109, 111, 116, 118, 123, 125]  # indexed to starting at 1 rather than zero
    test_labels = [3, 3, 9, 9, 17, 24, 24, 31, 31, 38, 38, 45, 45, 52, 52, 59, 59, 66, 66, 73, 73, 80, 80, 87, 87
                   , 94, 94, 101, 101, 108, 108, 115, 115, 122, 122]  # rows that hold the label of the measured burden

    labels = []
    for y in test_labels:
        label =  myburden.getblock([y, y+1, 3, 4])
        labels.append(label)

    burden_list = myburden.unique_list(labels)

    test_results = []
    for x in test_sets:
        aa = myburden.get_test_result(x)
        test_results.append(aa)
        print(aa, this_CT.burdenZ(aa - test_results[0], 5)) # complex burden, VA, PF

    # average for each burden, note merging of type-A and type-B
    print('Averaging')
    averaged_set = []
    for burden in burden_list:
        data_set = []
        for i in range(len(test_results)):
            if labels[i] == burden:
                data_set.append(test_results[i])

        average_typeA = gtc.type_a.estimate(data_set,label = 'type-A' + repr(burden))
        average_typeB = gtc.function.mean(data_set)
        averaged = gtc.type_a.merge_components(average_typeA, average_typeB)
        averaged_set.append(averaged)

    for i in range(len(burden_list)):
        print(burden_list[i], averaged_set[i])

    #  subtract nominal zero burden
    print('Zero correcting')
    corrected_bdns = []
    for i in range(len(burden_list)):
        #if i != 0:  # gtc printing gets upset with correlation subtracting identical ucomplex
        corrected_bdn = averaged_set[i] - averaged_set[0]
        corrected_bdns.append(corrected_bdn)
        if i != 0:  # polar maths doesn't work on zero
            print(burden_list[i], corrected_bdns[i], this_CT.burdenZgtc(corrected_bdns[i], 5.0))


    output_block = []
    header_row = ['Label', 'Resistance', 'uR', 'df', 'jR', 'ujR', 'df', '', '', '', '', '', '', '', '', '', '']
    output_block.append(header_row)
    assert len(test_results)==len(labels), "labels don't match results"
    for i in range(len(test_results)):
        myrow = []
        myrow.append(repr(labels[i][0]))
        myrow.append(float(test_results[i].real.x))
        myrow.append(float(test_results[i].real.u))
        myrow.append(float(test_results[i].real.df))
        myrow.append(float(test_results[i].imag.x))
        myrow.append(float(test_results[i].imag.u))
        myrow.append(float(test_results[i].imag.df))
        #need blanks to match maximum column width
        myrow.append('')
        myrow.append('')
        myrow.append('')
        myrow.append('')
        myrow.append('')
        myrow.append('')
        myrow.append('')
        myrow.append('')
        myrow.append('')
        myrow.append('')
        myrow.append('')
        output_block.append(myrow)
    header_row = ['Average and zero correct the results above', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
    output_block.append(header_row)
    header_row = ['Label', 'Resistance', 'uR', 'df', 'k', 'jR', 'ujR', 'df', 'k', 'VA', 'uVA', 'df', 'k', 'PF', 'uPF', 'df', 'k']
    output_block.append(header_row)
    for i in range(len(burden_list)):
        if i != 0:# avoid the zero minus zero entry
            myrow = []
            myrow.append(burden_list[i][0][0])
            myrow.append(float(corrected_bdns[i].real.x))
            myrow.append(float(corrected_bdns[i].real.u))
            myrow.append(float(corrected_bdns[i].real.df))
            myrow.append(gtc.reporting.k_factor(float(corrected_bdns[i].real.df)))
            myrow.append(float(corrected_bdns[i].imag.x))
            myrow.append(float(corrected_bdns[i].imag.u))
            myrow.append(float(corrected_bdns[i].imag.df))
            myrow.append(gtc.reporting.k_factor(float(corrected_bdns[i].imag.df)))
            va = this_CT.burdenZgtc(corrected_bdns[i], 5.0)  # at 5 A secondary
            x = va[0]
            y = va[1]
            myrow.append(float(x.x))
            myrow.append(float(x.u))
            myrow.append(float(x.df))
            myrow.append(gtc.reporting.k_factor(x.df))
            myrow.append(float(y.x))
            myrow.append(float(y.u))
            myrow.append(float(y.df))
            myrow.append(gtc.reporting.k_factor(y.df))
            output_block.append(myrow)
    myburden.data_store.makeworkbook(output_block, 'BurdenResults2')


    """

    """
