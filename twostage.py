from __future__ import division
from __future__ import print_function
import math
import GTC as gtc
from ExcelPython import CALCULATOR


class TWOSTAGE(object):
    """
    Encapsulates the analysis required for the scale buildup described in E056 using the two-stage current transformers.
    The aim is to replace the calculations done in Excel (CTCAL2.xls) and use GTC for the uncertainty calculation.

    Type B uncertainty terms are collected in the __init__ process. In general it is assumed that these are of a variable
    nature and should be created as new GTC uncertain numbers at each use. For example, subtracting two SR830 readings
    (i.e. the buildup calculation) would see the type B terms for gain disappear, or diminish, but in reality the
    readings may be taken on different ranges.

    Type A information is gathered from the experimental data.
    """

    # TODO no methods for impedance measurement data included.
    # TODO no method for equivalent circuit of CTs for estimating burden dependence etc.
    # TODO review methods for auditing clarity against the measurement equations.
    # TODO clarify nature of this class, i.e. should it hold all methods for a set of two-stage primary standards?
    # TODO create a CT2STAGE class for holding ratio and parameter information, with Ta, Tb and Tc (5:1) as instances
    # TODO additional uncertainties for input data such as meter and resistor calibration.

    def __init__(self):
        # taken from uncert_summary sheet of Ctcal32012.xlsm
        self.ct_x_stability = 0.017e-6  # Short-term stability of the real part of the CT ratio.
        self.df_ct_x_stability = 15
        self.ct_y_stability = 0.011e-6  # Short-term stability of the imaginary part of the CT ratio.
        self.df_ct_y_stability = 15
        self.sr830 = 0.015e-6  # SR830 calibration errors.
        self.df_sr830 = 10
        self.cmnmode = 0.1e-6  # micro volt common mode error.
        self.df_cmnmode = 5
        self.sp_mag = 0.5  # magnetic series-parallel correction.
        self.df_sp_mag = 5
        self.sp_cap = 0.2  # capacitive series-parallel correction.
        self.df_sp_cap = 5
        # from 'CTCAL2 repaired 2018plus2018Cal.xls'
        delta_r = 0.005  # assumed variation in burden
        r2 = 0.486  # main secondary burden (total with external leads)
        r1 = 0.126  # auxiliary secondary burden (total with external leads)
        z2 = 0.2699  # Ta main secondary leakage impedance
        z1 = 0.1248  # Ta auxiliary secondary leakage impedance
        self.brdn1a = delta_r / (r1 + z1)  # effect on the ratio of a 0.005 ohm burden uncertainty with zero auxiliary burden.
        self.df_brdn1a = 5
        self.brdn2a = delta_r / (r2 + z2)  # effect on the ratio 0.005 ohm burden uncertainty with 0.2 ohm auxiliary burden.
        self.df_brdn2a = 5
        z2 = 0.079  # Tb main secondary leakage impedance
        z1 = 0.037  # Tb auxiliary secondary leakage impedance
        self.brdn1b = delta_r / (r1 + z1)  # effect on the ratio of a 0.005 ohm burden uncertainty with zero auxiliary burden.
        self.df_brdn1b = 5
        self.brdn2b = delta_r / (r2 + z2)  # effect on the ratio 0.005 ohm burden uncertainty with 0.2 ohm auxiliary burden.
        self.df_brdn2b = 5
    def ishare(self, ratio, Rs, rls, vdrop, i2, target_i2, individual):
        """
        Calculates the share of current in each primary in parallel using measurements of voltage across
        each primary section when all are connected in series.

        :param ratio: list [k, m, n] to extract desrired ratios, k sections of m turns with a secondary of n turns
        :param Rs:  sedondary burden resistance
        :param rls:  secondary leakage impedance
        :param vdrop:  list of measured primary volt drops in ascending order
        :param i2:  list of secondary currents, measured as voltage over Rs matching the primary volt drops
        :param target_i2:  the i2 value used to normalise the volt drops, usually 1, i.e. 5 A through 0.2 ohm
        :param individual:  boolean when set to True the vdrops are measured across each individual winding
        :return:  list of current fractions in each primary in parallel; list sums to k
        """

        assert individual in [True, False], "'indvidual' not set to True or False in ishare"
        k = len(vdrop)  # number of sections
        assert k == ratio[0], 'data does not match expected number of sections in ishare'
        N = ratio[2]/(ratio[0]*ratio[1])  # km/n, the ratio for series connection
        vi = []  # normalised volt drops
        for i in range(k):
            vi.append(vdrop[i] / i2[i] * target_i2)

        V = []  # volt drop across each section
        V.append(vi[0])  # the first section is the normalised measured value
        for i in range(1, k, 1):
            if individual == True:
                V.append(vi[i])  # no difference between list V and vi
            elif individual == False:
                V.append(vi[i] - vi[i - 1])  # other sections by differences

        Vs = target_i2  # this is the voltage across Rs when volt drops were measured
        rlp = []  # the set of leakage impedance of each of the k primary turns
        for x in V:
            rlp.append(Rs * x / Vs / N - (Rs + rls) * (1 / N) ** 2 * 1 / k)

        glp = []  # the leakage conductance of each primary coil
        totalglp = 0
        for x in rlp:
            glp.append(1 / x)
            totalglp = totalglp + 1 / x

        current_proportion = []  # the fraction of current through each primary in parallel
        for x in glp:
            current_proportion.append(x / totalglp * k)
        print('cur prop = ', current_proportion)
        return current_proportion

    def couple(self, shuntv, xv, yv, ctratio_m, shunt, comr, uut_ratio):
        """
        Calculates mutual coupling between two primary sections connected in oppostion.
        All parameters are lists of the k-1 measurements made for all sections against the first.

        :param shuntv: voltage measured across the shunt on the secondary monitoring the primary current
        :param xv: real part of voltage across the common resistor
        :param yv: imaginary part of the voltage across the common resistor
        :param ctratio_m: ratio of monitoring transformer to give primary current from the shuntv values
        :param shunt: value of the secondary shunt
        :param comr: value of the common resistor
        :param uut_ratio: the ratio [k, n, m] for the transformer being tested
        :return: list of coupling coefficients
        """

        ctratio_sect = uut_ratio[2] / uut_ratio[1]  # the ratio for a single section
        k = len(xv) + 1  # one fewer measurements than sections
        coupling = []
        for i in range(k - 1):
            voltage = xv[i] + 1j * yv[i]  # complex voltage across common resistor
            isec = voltage / comr[i]  # current through common resistor
            itest = shuntv[i] / shunt[i] * ctratio_m[i]  # primary current
            iratio = isec / (itest / ctratio_sect)  # the difference current divided by full scale secondary current
            coupling.append(iratio)
        return coupling

    def magneticsp(self, couple, share, uut_ratio):
        """
        Calculates the series error minus the parallel error for a set of primary windings due to magnetic
        coupling differences between the windings. Note that the way coupling was measured forces the first couple
        element to be zero.

        :param couple: list of k-1 mutual couplings for k primary windings
        :param share: list of fractional share of current for k primary windings in parallel
        :param uut_ratio: list [k, m, n] of windings for transformer
        :return: two complex errors, no scaling to % or ppm
        """

        assert len(share) - len(couple) == 1, "sharing list should be one longer than coupling list"
        k = len(share)
        N = uut_ratio[2]/(uut_ratio[0] * uut_ratio[1])
        if k == 4:  # series/parallel available on Ta primaries
            x = (2-(share[0]+share[1]))/2  # x is offset required to make first two add to 2.0
            y = (2-(share[2]+share[3]))/2  # y is offset required to make last two add to 2.0
            sp_share = []
            sp_share.append(share[0] + x)
            sp_share.append(share[1] + x)
            sp_share.append(share[2] + y)
            sp_share.append(share[3] + y)

            sp_merror = 0  # summation for the serie/parallel magnetic error
            for i in range(1, k):  # the first element of share would have been multiplied by zero, so not needed
                sp_merror = couple[i - 1] * (sp_share[i]-1) + sp_merror
            sp_merror = N * sp_merror  # the sign of this is as derived in section A3 of E056.005, relies on aj defn
            sp_merror = sp_merror * gtc.ureal(1.0, self.sp_mag, self.df_sp_mag, label='s_sp_mag')  # type B factor
        else:
            sp_merror = 0  #i.e. there is no series/paralle connection

        merror = 0
        for i in range(1, k):  # the first element of share would have been multiplied by zero, so not needed
            merror = couple[i - 1] * (share[i] - 1) + merror
        merror = N * merror  #the sign of this is as derived in section A3 of E056.005 and depends on aj measurement
        merror = merror * gtc.ureal(1.0, self.sp_mag, self.df_sp_mag, label='sp_mag')  # type B factor

        return merror, sp_merror

    def oldcapsp(self, ypg, k, z4, r2, z2, uut_ratio):
        """
        Calculates the series error minus the parallel error for a set of primary windings due to the change
        in potential distribution on the capacitance from the windings to screen. INCORRECT old formula.

        :param ypg: ypg is complex admittance from series primary to primary screen
        :param k: the number of sections connected in series
        :param z4: the leakage impedance of the primary winding
        :param r2: the secondary burden impedance
        :param z2: leakage impedance of the secondary winding
        :param uut_ratio: [k, n, m] transformer ratio
        :return: error; should be a GTC ucomplex if inputs are experimental values
        """

        n = uut_ratio[2]/(uut_ratio[0] * uut_ratio[1])
        s_p_error = ypg / 6 * 1 / (k ** 2 - 1) * (z4 + (r2 + z2) / n ** 2)
        s_p_error = s_p_error * gtc.ureal(1.0, self.sp_cap, self.df_sp_cap, label = 'sp_cap')  # type B factor
        return s_p_error

    def newcapsp(self, ypg, z, r2, z2, uut_ratio):
        """
        Calculates the parallel error minus the series error for a set of primary windings due to the change
        in potential distribution across the distributed capacitance from the primary windings to the screen.
        This version of the formula differs from the original published version. It also caclualtes the series/
        parallel - series error when k=4. This error is set to 0 if k does not equal 4.

        :param ypg: ypg is complex admittance from series primary to primary screen
        :param z: is the leakage impedance of the primary when all the sections are connected in parallel
        :param r2: the secondary burden impedance
        :param z2: leakage impedance of the secondary winding
        :param uut_ratio: [k, m, n] is the ratio windings of the uut
        :return: (parallel-series) error, (series/parallel - series) erro
        """
        k = uut_ratio[0]
        m = uut_ratio[1]
        n = uut_ratio[2]
        s_p_error = ypg / 3 *(k**2 - 1) * ((r2 + z2) * (m/n) ** 2 + z)
        s_p_error = s_p_error * gtc.ureal(1.0, self.sp_cap, self.df_sp_cap, label='sp_cap')  # type B factor

        if k == 4:  # 4 sections in series/parallel possible, note relies on exact integer ... isclose() needed?
            g = 2
            sp_s_error = ypg / 3 * (k ** 2 - g**2) * ((r2 + z2) * (m / n) ** 2 + z)
            sp_s_error = sp_s_error * gtc.ureal(1.0, self.sp_cap, self.df_sp_cap, label='sp_s_cap')  # type B factor
        else:
            sp_s_error =0
        return s_p_error, sp_s_error


    def ratioerror(self, shuntv, xv, yv, shunt, comr, nomexcitation, full_scale):
        """
        Calculates the measured ratio error. All parameters are lists of data from

        :param shuntv: shuntv: voltage measured across the shunt on the secondary monitoring the primary current
        :param xv: xv: real part of voltage across the common resistor
        :param yv: imaginary part of the voltage across the common resistor
        :param shunt: value of the secondary shunt
        :param comr: value of the common resistor
        :param nomexcitation: nominal excitation list, e.g. 10%, 20% ... 120%
        :return: a list of errors, a list of actual excitation levels and a list of the matching nominal excitations
        """

        # TODO consider a 'swap' flag for where the nominal reference ratio was actually on the UUT side
        e = []  # list of error e1 at each excitation level
        excite = []  # matching list of fractional excitation level
        nom_excite = []  # matched nominal excitation level
        for i in range(len(xv)):
            xvolt = xv[i] * gtc.ureal(1.0, self.sr830, self.df_sr830, label = 'sr830 '+ repr(xv[i]))  #type B
            xvolt = xvolt + gtc.ureal(0.0, self.cmnmode, self.df_cmnmode, label = 'common mode ' + repr(xvolt.x))
            yvolt = yv[i] * gtc.ureal(1.0, self.sr830, self.df_sr830, label='sr830 '+ repr(yv[i]))  # type B
            yvolt = yvolt + gtc.ureal(0.0, self.cmnmode, self.df_cmnmode, label='common mode '+ repr(yvolt.x))
            ecurrent = (xvolt + 1j * yvolt) / comr[i]  # difference current measured by SR830
            current = shuntv[i] / shunt[i]  # primary =  secondary current
            e_answer =  gtc.result((ecurrent/current), 'my label')  # key intermiediate result? decide later
            e.append(e_answer)
            excite.append(current / full_scale * 100)  # note this is a ureal
            nom_excite.append(
                min(nomexcitation, key=lambda x: abs(x - excite[i].x)))  # finds closest nominal excitation level
        return e, excite, nom_excite

    def ctcompare(self, datablock, fullscale, targetexcitation, settings):
        """
        Takes a set of CT error measurements as a block and returns the calculated ratio errors. Older spread sheets
        did not include gain and reserve settings for the SR830, hence the boolean settings

        :param datablock: shunt Volts, X V, Y V, Sh V sd/100, X stdev V, Y stdev, blank, Shunt, Gain, Reserve, Com R.
        :param fullscale: the value of current, A, that equates to 100% excitation
        :param targetexcitation: the set of nominal % excitation levels at which measurements were made
        :param settings: a boolean set to True if gain and reserve colunmns available
        :return: a list of errors, a list of actual excitation levels and a list of the matching nominal excitations
        """

        assert settings in [True, False], "settings must be True for range and gain information, otherwise False"
        shuntv = []  # gtc values
        xv = []  # gtc values
        yv = []  # gtc values
        shunt = []  # secondary shunt duplicates of a fixed value
        gain = []  # SR830 range
        reserve = []  # reserve, 1 for on?
        comr = []  # common resistance, list will likely be duplicates of a fixed value

        for x in datablock:
            if settings == False:
                shunt.append(x[7])
                comr.append(x[8])
                shuntv.append(
                    gtc.ureal(x[0], x[3], 100, label='shuntv ' + str(x[0] / x[7] / fullscale)))  # use current as label
                xv.append(gtc.ureal(x[1], x[4], 100, label='X V ' + str(x[0] / x[7] / fullscale)))
                yv.append(gtc.ureal(x[2], x[5], 100, label='Y V ' + str(x[0] / x[7] / fullscale)))
            elif settings == True:
                gain.append(x[6])
                reserve.append(x[7])
                shunt.append(x[8])
                comr.append(x[9])
                shuntv.append(
                    gtc.ureal(x[0], x[3], 100, label='shuntv ' + str(x[0] / x[8] / fullscale)))  # use current as label
                xv.append(gtc.ureal(x[1], x[4], 100, label='X V ' + str(x[0] / x[8] / fullscale)))
                yv.append(gtc.ureal(x[2], x[5], 100, label='Y V ' + str(x[0] / x[8] / fullscale)))

        e, excite, nom_excite = self.ratioerror(shuntv, xv, yv, shunt, comr, targetexcitation, fullscale)
        return e, excite, nom_excite

    def signcheck(self, datablock, fullscale, settings):
        """
        The three measurements in datablock are gathered at equal excitation, 10%. The returned +/- 1.0 is
        multiplied by subsequently measured 'errors' to ensure that they are errors with respect to Ta
        sedondary. A smaller than ideal secondary current in Ta is a negative error. Note that the buildup alternates
        from Ta to Tb with corresponding sign changes. This method ensure a common convention in case the SR830 is
        reverse connected.

        :param datablock: three rows of data, normal, shunted with resistor and shunted with capacitor
        :param fullscale: the value of current in A that corresponds to full scale excitation
        :param settings: a boolean set to True if gain and reserve colunmns are available
        :return: either +1.0 or -1.0
        """
        e, excite, nom_excite = self.ctcompare(datablock, fullscale, [10, 10, 10], settings)  # assumed check at 10%
        if e[1].x.real - e[0].x.real > 0:
            signreal = -1.0
        else:
            signreal = 1.0

        if e[2].x.imag - e[0].x.imag > 0:
            signimag = -1.0
        else:
            signimag = 1.0
        assert signreal == signimag, 'error direction different for real and imaginary'
        sign = signreal
        return sign

    def magnetic_coupling(self, block1, block2, uut_ratio, Rs, rls, target_i2, individual):
        """
        Calculates the series error - parallel error for a set of primary windings.
        This method processes the data and calls on internal methods ishare, couple and magneticsp

        :param block1: the data block for calculating current sharing
        :param block2: the data block with coupling measurements
        :param uut_ratio: [k, m, n] ratio windings of the uut
        :param Rs:  secondary burden resistance
        :param rls:  secondary leakage impedance
        :param target_i2:  the i2 value used to normalise the volt drops, usually 1, i.e. 5 A through 0.2 ohm
        :param individual:  boolean when set to True the vdrops are measured across each individual winding
        :return: same as magneticsp, a complex error with no scaling
        """

        N = uut_ratio[2]/(uut_ratio[0] * uut_ratio[1])  # the ratio in series connection
        vdrop = []  # measured volt drop
        i2 = []  # measured secondary current voltage
        for x in block1:
            vdrop.append(x[2])
            i2.append(x[0])
        proportions = self.ishare(uut_ratio, Rs, rls, vdrop, i2, 1, individual)
        # data is label, shunt volts, X V, Y V, Sh V sd/100, X stdev V, Y stdev V, CT ratio, Shunt, Com R
        # k = len(block2) + 1  # one fewer measurements than sections
        # create lists of data
        shuntv = []  # gtc values
        xv = []  # gtc values
        yv = []  # gtc values
        ctratio = []  # duplicates of a constant?
        shunt = []  # secondary shunt duplicates of a fixed value
        comr = []  # common resistance duplicates of a fixed value

        for x in block2:
            shuntv.append(gtc.ureal(x[1], x[4], 100, label='shuntv ' + repr(x[0])))
            xv.append(gtc.ureal(x[2], x[5], 100, label='X V ' + repr(x[0])))
            yv.append(gtc.ureal(x[3], x[6], 100, label='Y V ' + repr(x[0])))
            ctratio.append(x[7])
            shunt.append(x[8])
            comr.append(x[9])

        # calculation of coupling
        coupling = self.couple(shuntv, xv, yv, ctratio, shunt, comr, uut_ratio)
        return self.magneticsp(coupling, proportions, uut_ratio)

    def interp(self, x0, x1, y0, y1, xx):
        """
        Simple two-point interpolation using (x0, y0) and (x1, y1) it interpolates value yy at point xx, i.e. (xx, yy)
        used for Tc calibration which is done against Ta but at a fifth of the calibrated excitation levels. The method
        is just in this class for convenience.

        :param x1: x value linked to y1
        :param y0: y value linked to x0
        :param y1: y value linked to x1
        :param xx: x value at which y value is required
        :return: yy required y value
        """
        yy = y0 + (y1 - y0) / (x1 - x0) * (xx - x0)
        return yy

    def one_fifth(self, e_orig, exc_orig):
        """
        For calibrating Tc, we need to estimate Ta corrections at lower excitations. The original errors at the original
        excitation levels are selected for interpolation to give the new errors at one_fifth excitation. Note that
        for Tc purposes these will be quoted as being the 1% to 125% errors at 1A.

        :param e_orig: array of errors at original excitation
        :param exc_orig:
        :return: e_fifth, an array of errors at one fifth the original excitation
        """
        ereal = []
        eimag = []
        for e in e_orig:
            ereal.append(e.real)
            eimag.append(e.imag)

        exc_new = []
        for exc in exc_orig:
            exc_new.append(exc/5.0)  # excitation level one fifth of normal

        e_fifth = []
        e_fifth.append(self.interp(exc_orig[4], exc_orig[5], ereal[4], ereal[5], exc_new[0])
                            +1j * self.interp(exc_orig[4], exc_orig[5], eimag[4], eimag[5], exc_new[0]))
        e_fifth.append(self.interp(exc_orig[4], exc_orig[5], ereal[4], ereal[5], exc_new[1])
                            +1j * self.interp(exc_orig[4], exc_orig[5], eimag[4], eimag[5], exc_new[1]))
        e_fifth.append(self.interp(exc_orig[5], exc_orig[6], ereal[5], ereal[6], exc_new[2])
                            +1j * self.interp(exc_orig[5], exc_orig[6], eimag[5], eimag[6], exc_new[2]))
        e_fifth.append(self.interp(exc_orig[5], exc_orig[6], ereal[5], ereal[6], exc_new[3])
                            +1j * self.interp(exc_orig[5], exc_orig[6], eimag[5], eimag[6], exc_new[3]))
        e_fifth.append(self.interp(exc_orig[6], exc_orig[7], ereal[6], ereal[7], exc_new[4])
                            +1j * self.interp(exc_orig[6], exc_orig[7], eimag[6], eimag[7], exc_new[4]))
        e_fifth.append(self.interp(exc_orig[7], exc_orig[8], ereal[7], ereal[8], exc_new[5])
                            +1j * self.interp(exc_orig[7], exc_orig[8], eimag[7], eimag[8], exc_new[5]))
        e_fifth.append(self.interp(exc_orig[7], exc_orig[8], ereal[7], ereal[8], exc_new[6])
                            +1j * self.interp(exc_orig[7], exc_orig[8], eimag[7], eimag[8], exc_new[6]))
        e_fifth.append(self.interp(exc_orig[7], exc_orig[8], ereal[7], ereal[8], exc_new[7])
                            +1j * self.interp(exc_orig[7], exc_orig[8], eimag[7], eimag[8], exc_new[7]))
        e_fifth.append(self.interp(exc_orig[7], exc_orig[8], ereal[7], ereal[8], exc_new[8])
                            +1j * self.interp(exc_orig[7], exc_orig[8], eimag[7], eimag[8], exc_new[8]))
        return e_fifth

    def extra_ratios(self, xp1ap, xp2ap, xp3as, xp2bs, mag1a, mag1asp, mag2a, mag2asp, mag3a, mag2b, cap1a, cap1asp, cap2asp, cap2b):
        """
        The transformer configurations used for the buildup are not exhaustive. Other connections are possible with
        errors that are calculated from the ratios determined in the buildup with magnetic and capacitive corrections
        specific to the connection. The errors calculated here are P3ap (500:5), P2as (25:5), P2asp (50:5), P1asp (10:5)
        and P2bp (200:5). The formulas are based on eqn 5b in the techproc.

        :param xp1ap: error list for primary 1 in parallel on Ta
        :param xp2ap:  error list for primary 2 in parallel on Ta
        :param xp3as:  error list for primary 3 in series on Ta
        :param xp2bs:  error list for primary 2 in series on Tb
        :param mag1a:  magnetic parallel error - series error for primary 1 on Ta
        :param mag1asp: magnetic series/parallel error - series error for primary 1 on Ta
        :param mag2a:   magnetic parallel error - series error for primary 2 on Ta
        :param mag2asp:  magnetic series/parallel error - series error for primary 2 on Ta
        :param mag3a:  magnetic parallel error - series error for primary 3 on Ta
        :param mag2b:  magnetic parallel error - series error for primary 2 on Tb
        :param cap1a:  capacitive parallel error - series error for primary 1 on Ta
        :param cap1asp:  capacitive series/parallel error - series error for primary 1 on Ta
        :param cap2asp:  capacitive series/parallel error - series error for primary 2 on Ta
        :param cap2b:  capacitive series/parallel error - series error for primary 2 on Tb
        :return:  the five lists of error for each of the five transformer connections
        """

        cap3a = 0
        p3ap = []
        for x in xp3as:
            p3ap.append(x  + mag3a + cap3a)

        cap2a = 0
        p2as = []
        for x in xp2ap:
            p2as.append(x - mag2a - cap2a)

        p2asp = []
        for x in xp2ap:
            p2asp.append(x + mag2asp - mag2a + cap2asp - cap2a)

        p1asp = []
        for x in xp1ap:
            p1asp.append(x + mag1asp - mag1a + cap1asp -cap1a)

        p2bp = []
        for x in xp2bs:
            p2bp.append(x + mag2b + cap2b)

        return p3ap, p2as, p2asp, p1asp, p2bp

    def buildup(self, excite, e1, e2, e3, e4, e5a, e5b, mag1a, mag2a, mag3a, mag1b, mag2b, capa, capb):
        """
        Assembles all the measurement results so that final ratio errors can be calculated.
        It is assumed that all input error values are signed as an error in Ta after using the signcheck method

        :param excite: the list of % excitation levels at which the errors apply
        :param e1: Ta 5:5 self cal of P1as
        :param e2: 20:5 P1ap vs P1bs
        :param e3: 100:5 P1bp vs P1ap
        :param e4: 100:5 P2ap vs P2bs
        :param e5: 100:5 P1bp vs P3as
        :param mag1a: magnetic series parallel error for P1a
        :param mag2a: magnetic series parallel error for P2a
        :param mag3a: magnetic series parallel error for P3a
        :param mag1b: magnetic series parallel error for P1b
        :param mag2b: magnetic series parallel error for P2b
        :param capa: capacitive series parallel erro for P1a
        :param capb: capacitive series parallel error for P1b
        :return: list of errors for p1as, p1ap, p2ap, p3as, p1bs, p1bp, p2bs
        """

        # TODO should the measured errors already be fully corrected, or should some corrections occur in this method?
        #P1as, 5:5 in series

        polarity = 1  # SR830 readings normalised to be 'error' when Ta is the UUT
        p1as = []
        for x in e1:
            ans = x * polarity
            ans = ans + gtc.ucomplex(0 + 0j,(self.ct_x_stability, self.ct_y_stability),self.df_ct_x_stability,
                                     label = 'ct_stability_p1as')
            ans = ans * gtc.ureal(1.0, self.brdn1a, self.df_brdn1a, label="burden1a")
            ans = ans * gtc.ureal(1.0, self.brdn2a, self.df_brdn2a, label="burden2a")
            p1as.append(ans)  # as measured, modified by polarity

        p1ap = []
        for x in p1as:
            ans = x + mag1a + capa
            ans = ans + gtc.ucomplex(0 + 0j,(self.ct_x_stability, self.ct_y_stability),self.df_ct_x_stability,
                                     label = 'ct_stability_p1ap')
            ans = ans * gtc.ureal(1.0, self.brdn1b, self.df_brdn1b, label = "burden1b")
            ans = ans * gtc.ureal(1.0, self.brdn2b, self.df_brdn2b, label="burden2b")
            p1ap.append(ans)

        polarity = -1 # as Tb is being calibrated
        p1bs = []
        for i in range(len(p1ap)):
            ans = p1ap[i] + e2[i] * polarity
            ans = ans + gtc.ucomplex(0 + 0j,(self.ct_x_stability, self.ct_y_stability),self.df_ct_x_stability,
                                     label = 'ct_stability_p1bs')
            p1bs.append(ans)

        p1bp = []
        for i in range(len(p1bs)):
            ans = p1bs[i] + capb + mag1b
            ans = ans + gtc.ucomplex(0 + 0j,(self.ct_x_stability, self.ct_y_stability),self.df_ct_x_stability,
                                     label = 'ct_stability_p1bp')
            p1bp.append(ans)

        polarity = 1  # as Ta is being calibrated
        p2ap = []
        for i in range(len(p1bp)):
            ans = p1bp[i] + e3[i] * polarity
            ans = ans + gtc.ucomplex(0 + 0j,(self.ct_x_stability, self.ct_y_stability),self.df_ct_x_stability,
                                     label = 'ct_stability_p2ap')
            p2ap.append(ans)

        polarity = -1  # as Tb is being calibrated
        p2bs = []
        for i in range(len(p2ap)):
            ans = p2ap[i] + e4[i] * polarity
            ans = ans + gtc.ucomplex(0 + 0j,(self.ct_x_stability, self.ct_y_stability),self.df_ct_x_stability,
                                     label = 'ct_stability_p2bs')
            p2bs.append(ans)

        polarity = 1  # as Ta is being calibrated
        p3as_a = []
        for i in range(len(p2bs)):
            ans = p2bs[i] + e5a[i] * polarity
            ans = ans + gtc.ucomplex(0 + 0j,(self.ct_x_stability, self.ct_y_stability),self.df_ct_x_stability,
                                     label = 'ct_stability_p3as_a')
            p3as_a.append(ans)


        polarity = 1  # as Ta is being calibrated
        p3as_b = []
        for i in range(len(p1bp)):
            ans = p1bp[i] + e5b[i] * polarity
            ans = ans + gtc.ucomplex(0 + 0j,(self.ct_x_stability, self.ct_y_stability),self.df_ct_x_stability,
                                     label = 'ct_stability_p3as_b')
            p3as_b.append(ans)

        return p1as, p1ap, p2ap, p3as_a, p3as_b, p1bs, p1bp, p2bs


if __name__ == "__main__":
    # TODO check for and record agreement/disagreement with old Excel calculations
    # The Excel workbooks collect data in a 'Data(xxxx)' sheet where each row has an index.
    # This is the likely input data form that python will use.

    labdata = CALCULATOR("Test.xlsm", "TwoStageResults.xlsx")
    calrun = TWOSTAGE()

    # Work through the buildup process to sort out an efficient set of methods

    target_excitation = [120, 100, 60, 40, 20, 10,5]  # buildup needs a common set of nominal excitation levels

    # First 5:5 ratio for T_a using the four inner primary windings, primary 1, connected in series
    print('Step1')
    # block_descriptor = [57, 63, 4, 13]  # Ta
    block_descriptor = [155, 162, 4, 14]  # Ta
    copy_data1 = labdata.getdata_block('Data2013b', block_descriptor)
    full_scale = 5.0  # nominal 100 % excitation current
    e1, excite1, nom_excite1 = calrun.ctcompare(copy_data1, full_scale, target_excitation, True)

    # Second step 20:5 ratio from Ta, primary 1 in parallel, to Tb, primary 1 in series
    print('Step2')
    # block_descriptor = [112, 118, 4, 13]
    block_descriptor = [169, 177, 4, 14]
    copy_data2 = labdata.getdata_block('Data2013b', block_descriptor)
    full_scale = 5.0  # nominal 100 % excitation current
    e2, excite2, nom_excite2 = calrun.ctcompare(copy_data2, full_scale, target_excitation, True)

    # Third step 100:5 ratio from Tb, primary 1 in parallel, to Ta, primary 2 in parallel
    print('Step 3')
    # block_descriptor = [164, 170, 4, 13]
    block_descriptor = [184, 191, 4, 14]
    copy_data3 = labdata.getdata_block('Data2013b', block_descriptor)
    full_scale = 5.0  # nominal 100 % excitation current
    e3, excite3, nom_excite3 = calrun.ctcompare(copy_data3, full_scale, target_excitation, True)

    # Fourth step 100:5 from Ta, primary 2 in parallel, to Tb, primary 2 in series
    print('Step 4')
    # block_descriptor = [244, 250, 4, 13]
    block_descriptor = [248, 255, 4, 14]
    copy_data4 = labdata.getdata_block('Data2013b', block_descriptor)
    full_scale = 5.0  # nominal 100 % excitation current
    e4, excite4, nom_excite4 = calrun.ctcompare(copy_data4, full_scale, target_excitation, True)

    # Fifth step  100:5 from Tb, primary 1 in parallel (or primary 2 in series), to Ta ,primary 3 in series
    print('Step 5')
    # block_descriptor = [172, 178, 4, 13]
    block_descriptor = [194, 201, 4, 14]
    copy_data5 = labdata.getdata_block('Data2013b', block_descriptor)
    full_scale = 5.0  # nominal 100 % excitation current
    e5, excite5, nom_excite5 = calrun.ctcompare(copy_data5, full_scale, target_excitation, True)
    print('e5 = ', e5)

    # get coupling measurements for Ta primary 1
    print('Coupling measurements 1a')
    block_descriptor = [84, 87, 3, 13]  # coupling results for inner primary of Ta
    copy_dataA = labdata.getdata_block('Data2013', block_descriptor)
    # get volt drops of inner primary Ta
    ratio = 100 / 100  # N2/N1, Ta primary 1
    Rs = 0.2  # resistance of secondary burden
    rls = 0.268  # leakage resistance of Ta secondary, measured earlier
    Vs = 5 * Rs  # nominal voltage across burden at 5 A
    block_descriptor = [51, 55, 4, 8]
    copy_dataB = labdata.getdata_block('Data2013', block_descriptor)
    magnetic1a = calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, False)

    # get coupling measurements for Tb primary 1
    print('Coupling measurements 1b')
    block_descriptor = [139, 143, 3, 13]  # coupling results for inner primary of Tb
    copy_dataA = labdata.getdata_block('Data2013', block_descriptor)
    # get volt drops of middle primary Tb
    ratio = 30 / 120  # N2/N1, Tb primary 1
    Rs = 0.2  # resistance of secondary burden
    rls = 0.268  # leakage resistance of Tb secondary, measured earlier
    Vs = 5 * Rs  # nominal voltage across burden at 5 A
    block_descriptor = [127, 132, 4, 8]
    copy_dataB = labdata.getdata_block('Data2013', block_descriptor)
    magnetic1b = calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, False)

    # get coupling measurements for Ta primary 2
    print('Coupling measurements 2a')
    block_descriptor = [150, 153, 3, 13]  # coupling results for middle primary of Ta
    copy_dataA = labdata.getdata_block('Data2013', block_descriptor)
    # get volt drops of middle primary Ta
    ratio = 20 / 100  # N2/N1, Ta primary 2
    Rs = 0.2  # resistance of secondary burden
    rls = 0.268  # leakage resistance of Ta secondary, measured earlier
    Vs = 5 * Rs  # nominal voltage across burden at 5 A
    block_descriptor = [157, 161, 4, 8]
    copy_dataB = labdata.getdata_block('Data2013', block_descriptor)
    magnetic2a = calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, False)

    # get coupling measurements for Ta primary 3
    print('Coupling measurements 3a')
    block_descriptor = [356, 360, 3, 13]  # coupling results for outer primary of Ta
    copy_dataA = labdata.getdata_block('Data2013', block_descriptor)
    # get volt drops of outer primary Ta
    ratio = 5 / 100  # N2/N1, Ta primary 3
    Rs = 0.2  # resistance of secondary burden
    rls = 0.268  # leakage resistance of Ta secondary, measured earlier
    Vs = 5 * Rs  # nominal voltage across burden at 5 A
    block_descriptor = [190, 195, 4, 8]
    copy_dataB = labdata.getdata_block('Data2013', block_descriptor)
    magnetic3a = calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, False)

    # get coupling measurements for Tb primary 2
    print('Coupling measurements 2b')
    block_descriptor = [226, 227, 3, 13]  # coupling results for outer primary of Tb
    copy_dataA = labdata.getdata_block('Data2013', block_descriptor)
    # get volt drops of outer primary Tb
    ratio = 6 / 120  # N2/N1, Tb primary 2
    Rs = 0.2  # resistance of secondary burden
    rls = 0.268  # leakage resistance of Tb secondary, measured earlier
    Vs = 5 * Rs  # nominal voltage across burden at 5 A
    block_descriptor = [222, 224, 4, 8]
    copy_dataB = labdata.getdata_block('Data2013', block_descriptor)
    # print(len(copy_dataA), len(copy_dataB))
    magnetic2b = calrun.magnetic_coupling(copy_dataB, copy_dataA, ratio, Rs, rls, 1, False)

    frequency = 50.0  # Hz, the frequency of calibration
    capacitance = 725.3e-12  # farad, capacitance from 100 turn primary to screen
    ypg_admit = 1j * 2 * math.pi * frequency * capacitance
    k = 4  # number of sections on the primary
    z4 = 0.350 + 1j * 0.033  # ohm, primary leakage impedance from Y51, Z51 of 'Calc 50 Hz' in CTCAL2
    z3 = 0.266 + 1j * 0.002  # ohm, secondary leakage impedance from V19, W19 of 'Calc 50 Hz' in CTCAL2
    r2 = 0.2  # ohm, secondary burden impedance
    ratio = 100 / 25  # ratio of series connection
    cap_errora = calrun.newcapsp(ypg_admit, k, z4, r2, z3, ratio)
    print('capacitance error = ', cap_errora)
    # TODO check result of the capsp method matches Andy's results...No!

    frequency = 50.0  # Hz, the frequency of calibration
    capacitance = 725.3e-12  # farad, capacitance from 100 turn primary to screen
    ypg_admit = 1j * 2 * math.pi * frequency * capacitance
    k = 4  # number of sections on the primary
    z4 = 0.350 + 1j * 0.033  # ohm, primary leakage impedance from Y51, Z51 of 'Calc 50 Hz' in CTCAL2
    z3 = 0.266 + 1j * 0.002  # ohm, secondary leakage impedance from V19, W19 of 'Calc 50 Hz' in CTCAL2
    r2 = 0.2  # ohm, secondary burden impedance
    ratio = 100 / 25  # ratio of series connection
    cap_errorb = calrun.newcapsp(ypg_admit, k, z4, r2, z3, ratio)
    print('capacitance error = ', cap_errorb)

    # Use results above to calculate ratios for each transformer
    t1, t2, t3, t4, t5, t6, t7 = calrun.buildup(target_excitation, e1, e2, e3, e4, e5, magnetic1a, magnetic2a, magnetic3a, magnetic1b, magnetic2b, cap_errora, cap_errorb)


    # Now everything needs to be manipulated into tables for presenting in a spreadsheet
    theblock = []
    theblock.append(['', 'e1.real', 'e2.real', 'e3.real', 'e4.real', 'e5.real', ''])
    for i in range(len(e1)):
        output = []
        output.append(nom_excite1[i])
        output.append(e1[i].x.real)
        output.append(e2[i].x.real)
        output.append(e3[i].x.real)
        output.append(e4[i].x.real)
        output.append(e5[i].x.real)
        output.append('') # to match width
        theblock.append(output)

    theblock.append(['', 'e1.imag', 'e2.imag', 'e3.imag', 'e4.imag', 'e5.imag', ''])
    for i in range(len(e1)):
        output = []
        output.append(nom_excite1[i])
        output.append(e1[i].x.imag)
        output.append(e2[i].x.imag)
        output.append(e3[i].x.imag)
        output.append(e4[i].x.imag)
        output.append(e5[i].x.imag)
        output.append('')  # to match width
        theblock.append(output)

    theblock.append(['', '1a', '1b', '2a', '3a', '2b', ''])
    theblock.append(
        ['real', repr(magnetic1a.x.real), repr(magnetic1b.x.real), repr(magnetic2a.x.real), repr(magnetic3a.x.real), repr(magnetic2b.x.real), ''])
    theblock.append(
        ['imag', repr(magnetic1a.x.imag), repr(magnetic1b.x.imag), repr(magnetic2a.x.imag), repr(magnetic3a.x.imag), repr(magnetic2b.x.imag), ''])
    theblock.append(['', 'real', 'imag', '', '', '',''])
    theblock.append(['capa s - p', repr(cap_errora.real), repr(cap_errora.imag), '', '', '', ''])
    theblock.append(['capb s - p', repr(cap_errorb.real), repr(cap_errorb.imag), '', '', '', ''])

    tableheader = ['Excitation', 'Real', 'Imag','U real', 'U imag', 'k real','k imag']
    theblock.append(['','','','','','',''])  # space
    theblock.append(['P1as','','','','','',''])  # title
    theblock.append(tableheader)
    block_t1 = labdata.ucomplex_table(target_excitation, t1)
    for row in block_t1:
        theblock.append(row)

    theblock.append(['','','','','','',''])  # space
    theblock.append(['P1ap','','','','','',''])  # title
    theblock.append(tableheader)
    block_t2 = labdata.ucomplex_table(target_excitation, t2)
    for row in block_t2:
        theblock.append(row)

    theblock.append(['','','','','','',''])  # space
    theblock.append(['P2ap','','','','','',''])  # title
    theblock.append(tableheader)
    block_t3 = labdata.ucomplex_table(target_excitation, t3)
    for row in block_t3:
        theblock.append(row)

    theblock.append(['','','','','','',''])  # space
    theblock.append(['P3as','','','','','',''])  # title
    theblock.append(tableheader)
    block_t4 = labdata.ucomplex_table(target_excitation, t4)
    for row in block_t4:
        theblock.append(row)

    theblock.append(['','','','','','',''])  # space
    theblock.append(['P1bs','','','','','',''])  # title
    theblock.append(tableheader)
    block_t5 = labdata.ucomplex_table(target_excitation, t5)
    for row in block_t5:
        theblock.append(row)

    theblock.append(['','','','','','',''])  # space
    theblock.append(['P1bp','','','','','',''])  # title
    theblock.append(tableheader)
    block_t6 = labdata.ucomplex_table(target_excitation, t6)
    for row in block_t6:
        theblock.append(row)

    theblock.append(['','','','','','',''])  # space
    theblock.append(['P2bs','','','','','',''])  # title
    theblock.append(tableheader)
    block_t7 = labdata.ucomplex_table(target_excitation, t7)
    for row in block_t7:
        theblock.append(row)

    labdata.makeworkbook(theblock, 'my_sheet_name')

    print()
    print('Example budget, t7[3]')
    print()
    print(repr(t7[3].real))
    for l, u in gtc.rp.budget(t7[3].real, trim = 0):
        print(l, u)

    print()
    print('Check new capacitance error.')
    print()
    k = 4  # sections of
    m = 25  # turns
    n = 100  # turns secondary
    r = 1  # ohm burden
    ypg = 1j*3.14159*2*50*725e-12  # 725 pF admittance of primary to screen
    z2=0
    z=0
    print(repr(calrun.newcapsp(ypg,k,z,r,z2,m,n)))
