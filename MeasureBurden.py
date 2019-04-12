from __future__ import division
"""
Calculating value of burden by perturbing secondary current with a resistor
across output terminals of a CT
"""
import ExcelPython
import modelCT  #just for burden
import arnold  # uncertainty terms for dials
import GTC as gtc

class BURDEN(object):
    """
    Simple measurement method based on perturbing the burden by connecting an impedance in parallel.
    Changes in the secondary current error of the CT are then used to calculate the burden.
    """
    def __init__(self, excelfile, datasheet):
        """
        'excelfile' contains the worksheet 'datasheet' that contains all the essential data
        """
        self.excelfile = excelfile
        self.datasheet = datasheet
        self.data_store = ExcelPython.CALCULATOR(self.excelfile, "myBurdenResults.xlsx")
        self.arnie = arnold.ARNOLD('source', 'output)')  # dummy source and output as not needed

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

    def get_test_result(self, a):
        """
        :param a: first row of test
        :return: list of (ultimately ureals) [complex error 1, complex error 2, resistance]
        """
        # first get relevant resistor
        resistor = float(self.data_store.getdata_block(self.datasheet, [a +1, a + 3 , 12, 13])[0][0])
        multiplier = float(self.data_store.getdata_block(self.datasheet, [a , a + 1 , 7, 8])[0][0])

        errors = self.data_store.getdata_block(self.datasheet, [a , a + 2 , 10, 12])
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
        return the_burden


if __name__ == "__main__":
    #use fitted impedance coefficients to define behviour of core impedance
    this_core = [11.7204144319303, -0.0273771774408702, 29.7364068940134, 3.90272331971033, 0.000243735060998679, 5.8544055409592]
    this_CT = modelCT.CT(this_core)  # a convenience CT
    #analyse data from Excel sheet
    myburden = BURDEN('S22008_py_Burden_5AmpRatio Template V5.0_draft.xlsm', 'ForPython')
    print('starting run')
    test_sets = [4, 6, 11, 13, 18, 25, 27, 32, 34, 39, 41, 46, 48, 53, 55, 60, 62, 67, 69,74, 76, 81, 83, 88,
                 90, 95, 97, 102, 104, 109, 111, 116, 118, 123, 125]
    test_labels = [3, 3, 9, 9, 17, 24, 24, 31, 31, 38, 38, 45, 45, 52, 52, 59, 59, 66, 66, 73, 73, 80, 80, 87, 87
                   , 94, 94, 101, 101, 108, 108, 115, 115, 122, 122]  # rows that hold the label of the measured burden
    labels = []
    for y in test_labels:
        label =  myburden.data_store.getdata_block('ForPython', [y, y+1, 3, 4])
        labels.append(label)
    test_results = []
    for x in test_sets:
        aa = myburden.get_test_result(x)
        test_results.append(aa)
        print(aa, this_CT.burdenZ(aa - test_results[0], 5)) # complex burden, VA, PF

    output_block = []
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
        output_block.append(myrow)

    myburden.data_store.makeworkbook(output_block, 'BurdenResults')

    """
C:\Python27\python.exe Y:/Staff/KJ/PycharmProjects/CTs/MeasureBurden.py
(-0.0074+0.0173078484156j)
(0.0220004818951+0.00145450504084j)
starting run
((0.02200048189510987+0.0014545050408433385j), (0.0, 1.0))
((0.021704696374683675+0.0011640579808546538j), (0.01036365550704316, 0.7135163847958342))
((0.07100500750434612+0.005818590383620416j), (1.2299615774144461, 0.9960580580137046))
((0.07185141115242598+0.004223942421912533j), (1.248194929494588, 0.9984604183078494))
((0.07145082523116554+0.004369560626657133j), (1.2384047210832798, 0.9982670142923787))
((0.07185142323620858+0.004078289235442487j), (1.2479985466972439, 0.9986177762992262))
((0.07245228936983428+0.0040783382105181185j), (1.26299974712214, 0.9986503874938109))
((0.12174759636821383+0.0067067296187779245j), (2.497132444015481, 0.9986165803114845))
((0.12315113660005982+0.005832102301849859j), (2.531133430034308, 0.9990648211656993))
((0.1710917356273901+0.007005226391091905j), (3.7298636408529804, 0.9993076697180857))
((0.17269706405609542+0.008173026090710021j), (3.771156851691879, 0.99900765260783))
((0.18780955633185326+0.1197127361461439j), (5.091516129167973, 0.8141439122959185))
((0.18660506087059658+0.11970985769641886j), (5.066988494867524, 0.8121420600333842))
((0.27423841690417294+0.1772469793753108j), (7.686309659480236, 0.8204130011141129))
((0.27343403203897115+0.17724413568058975j), (7.669779347315184, 0.8195592687809066))
((0.1712921454139395+0.008756568050940147j), (3.7367533468593157, 0.9988059798240833))
((0.17289770735988197+0.008464953793823194j), (3.776499623054101, 0.9989225508166455))
((0.022204915946775077+0.001164069624048356j), (0.008879259992009739, -0.5755942833332188))
((0.02220489901230924+0.001746104435776912j), (0.008902828375357996, 0.5740229637728758))
((0.022007373257484822+0.0030970018143483234j), (0.04106278075893111, 0.004195625726986797))
((0.022088096667782928+0.00242504983025557j), (0.02436228561534437, 0.089908202843121))
((0.07184112291372553+0.0036590268178321498j), (1.2472342937003988, 0.9990232242320786))
((0.07222649660395687+0.0038216148401385764j), (1.2570440928859743, 0.9988912678778041))
((0.02230992312444059+0.0011645914179211304j), (0.010600819047530786, -0.7297578327280178))
((0.02230992312444059+0.0011645914179211304j), (0.010600819047530786, -0.7297578327280178))
((0.0717523866921482+0.004667578994785433j), (1.2463887523479247, 0.9979210880898232))
((0.0721535250609397+0.004740585873198771j), (1.2565145408108005, 0.9978603815732047))
((0.12134287150464967+0.006722941758150194j), (2.4870498083545027, 0.9985967035705181))
((0.12274955075073206+0.007307955598369649j), (2.5229741574558764, 0.998316496404543))
((0.17203843690734727+0.008640355868023083j), (3.7552483603255733, 0.9988550730585325))
((0.17294482860195762+0.008128086027331453j), (3.7772950379488064, 0.9990240713948535))
((0.18772630564988435+0.11811052551314369j), (5.066660375165303, 0.8177271194606545))
((0.18757448832897783+0.11825635079734746j), (5.065657828622109, 0.8171397083826779))
((0.2755093806451681+0.17469913567139886j), (7.676280991949523, 0.8256240848137429))
((0.27505645956560265+0.17440187992072526j), (7.662739486983887, 0.8256054446987912))

Process finished with exit code 0
    """
