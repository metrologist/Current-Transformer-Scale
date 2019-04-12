from __future__ import division
from __future__ import print_function

class TRANSFORMER(object):
    """
    The TRANSFORMER class captures the essential structure of a transformer. It is specifically constructed for use
    with MSL's two-stage primary reference current transformers. These transformers have a main secondary and a second
    core auxiliary secondary. Primary windings are around both cores layered in groups with each group having a series
    or parallel connection. It should also be useable for a single core transformer with multiple secondary taps and
    fixed or window wound primary windings.

    There will likely be some experimentation to find the optimum dictionary structure to best describe a transformer.
    The calibration procedure should provide a dictionary of calibration constants with keys that match the ratios.

    :param short_label: short label to identify transformer
    :param primaries: a dictionary of sets of primary windings
    :param secondaries: a dictionary of sets of secondary windings
    :param cores: a dictionary of cores (either 1 or 2 cores)
    :param trans_type: either current or voltage
    """
    #TODO Describe primary as k sections of m turns, rather than just a list of turns.

    def __init__(self, short_label, primaries, secondaries, cores, trans_type):
        self.short_label = short_label
        self.primaries = primaries
        self.secondaries = secondaries
        self.cores = cores
        assert trans_type in ['voltage', 'current'], "transformer type must be voltage or current"
        self.trans_type = trans_type
        self.calibrated = False  # instantiated with no calibration data by default

    def nominal_ratio(self, primary, secondary):
        if self.trans_type == 'voltage':
            rat = primary / secondary
        elif self.trans_type == 'current':
            rat = secondary / primary
        else:
            rat = 0  # note that assertion ensures the transformer type exists so don't expect this result
        return rat

    def series(self, primary):
        return sum(primary)

    def parallel(self, primary):
        for i in range(1, len(primary)):
            assert primary[i] == primary[0], 'parallel windings must all have identical number of turns'
        return primary[0]

    def series_parallel(self, primary):
        for i in range(1, len(primary)):
            assert primary[i] == primary[0], 'parallel windings must all have identical number of turns'
        # should generalise so that 4 sections go 2x2, 6 sections 3x2 or 2x3, 8 sections 4x2 or 2x4 ?
        # for now, just 2x2, 3x2, 4x2
        return sum(primary)/2

    def dictionary_of_ratios(self):
        """
        create a list of the ratios for all series and all parallel connections
        designed for our 2-stage primary references
        in principle should be extendable to multi-ratio transformers in general
        probably will only work with dictionaries??
        :return: a dictionary with a key = primary label + short label + 's' or 'p' for series or parallel
        """
        ratios_dict = {}
        for x in self.primaries:
            ratios_dict[x + self.short_label + 's'] = self.nominal_ratio(self.series(self.primaries[x]),
                                                                         self.secondaries['main'])
            ratios_dict[x + self.short_label + 'p'] = self.nominal_ratio(self.parallel(self.primaries[x]),
                                                                         self.secondaries['main'])
            if len(self.primaries[x])%2==0 and len(self.primaries[x])>2:  # check if even and 4 or more
                ratios_dict[x + self.short_label + 'sp'] = self.nominal_ratio(self.series_parallel(self.primaries[x]),
                                                                              self.secondaries['main'])
        return ratios_dict

    def set_calibration(self, cal_errors):
        """
        Run set_calibration to pick up calibrated errors from a dictionary. These errors are available for when
        the transformer is used for calibrating others. Could be part of the __init__ process. TRANSFORMER could
        return errors interpolated for different burdens or excitation levels.

        :param cal_errors: dictionary of calibrated errors, i.e. the calibration certificate
        :return: sets calibration values for this instance and sets the 'calibrated' flag to True
        """
        self.cal = cal_errors
        self.calibrated = True
        return

    def make_caldict(self, input):
        # perhaps take data from file to contstruct the dictionary
        # use set_calibration to make it available in the object
        # have method to extract each complex error from the dictionary
        # initially not adding much value, but this method could incorporate interpolation for exciation and burdens
        # the errors should be as in a cal cert, i.e. not normally needing furnter processing
        # use GTC archive for filing data and, probably, gtc.result() to identify key intermediate calculations
        return

    def check_caldict(self, cal):
        """
        Checks that calibration information is consistent with the instantiated transformer description

        :param cal: calibration dictionary {'Ref' : , 'Exc' : , 'Isec' : , 'Bdn' :  'Named errors' : ... }
        :return: boolean, True if consistent, otherwise False
        """
        ok = True  # until proven false
        ratio_names = self.dictionary_of_ratios()
        for z in ratio_names:
            if z not in cal:  # check there is an error entry for every ratio
                ok = False
                print(z, 'ratio not in cal dictionary')

        compulsory = ['Ref', 'Exc', 'Isec', 'Bdn']
        for y in compulsory:  # check essential dictionary keys are provided
            if y not in cal:
                ok = False
                print(y, 'key not in cal dictionary')

        assert ok, "cal dictionary already failed basic items, not able to continue tests"

        levels = len(cal['Exc'])
        no_of_burdens = len(cal['Bdn'])
        for z in ratio_names:
            if len(cal[z]) != no_of_burdens:  # check there is an error set for each burden
                ok = False
                print(z, 'does not have correct number of error sets for the number of burdens')
                print(len(cal[z]))
            assert ok, "cal dictionary already failed correct error sets, not able to continue tests"
            for i in range(no_of_burdens):
                if len(cal[z][i]) != levels:  # check there is an error for each excitation level
                    ok = False
                    print(z, 'does not have the correct number of errors to match the excitation level')
        return ok


if __name__ == "__main__":
    print("Welcome to TRANSFORMER")
    # construct transformers using lists only
    ta_primaries = {'P1': [25, 25, 25, 25], 'P2': [5, 5, 5, 5], 'P3': [1, 1, 1, 1, 1]}
    ta_secondaries = {'main': 100, 'auxiliary': 100}
    ta_cores = {'main': 'could be a polynomial', 'auxiliary': 'could be another polynomial'}
    Ta = TRANSFORMER('a', ta_primaries, ta_secondaries, ta_cores, 'current')
    print()
    print('listing ratios')
    print('For Ta')
    answer = Ta.dictionary_of_ratios()
    print(answer)

    # construct the same transformers using dictionaries
    # this means that the windings can be referenced by name rather than just a list integer index

    tb_primaries = {'P1': [6, 6, 6, 6], 'P2': [3, 3]}
    tb_secondaries = {'main': 120, 'auxiliary': 120}
    tb_cores = {'main': 'could be a polynomial', 'auxiliary': 'could be another polynomial'}
    Tb = TRANSFORMER('b', tb_primaries, tb_secondaries, tb_cores, 'current')

    print()
    print('listing ratios')

    print('For Tb')
    answer = Tb.dictionary_of_ratios()
    print(answer)

    # how should calibration errors be incorporated?
    # expect that the class is used without calibration to assist in setting ratios and names for calibration
    # calibration instantiated from file
    # file needs to have date and/or calibration report reference ... basis for all our calibration reports?

    reference = "21 June 2018, S22xxx"  # the cal date and report number or other reference
    excitation = [1, 5, 10, 20, 40, 60, 80, 100, 120]
    i2100 = 5  # the current in amps at 100 % excitation
    burdens = [1]  # list of complex resistance of burdens at which the transformer was calibrated
    eP1ap = [[1, 5, 10, 20, 40, 60, 80, 100, 120]]  # error of P1ap connected to burden 1, as list of complex errors
    eP1as = [[1, 5, 10, 20, 40, 60, 80, 100, 120]]  # perhaps a list of lists, one list for each burden
    eP2ap = [[1, 5, 10, 20, 40, 60, 80, 100, 120]]
    eP2as = [[1, 5, 10, 20, 40, 60, 80, 100, 120]]
    eP3ap = [[1, 5, 10, 20, 40, 60, 80, 100, 120]]
    eP3as = [[1, 5, 10, 20, 40, 60, 80, 100, 120]]  # error of P3as connected to burden 1
    eP2asp = [[1, 5, 10, 20, 40, 60, 80, 100, 120]]
    eP1asp = [[1, 5, 10, 20, 40, 60, 80, 100, 120]]
    caldict = {'Ref': reference, 'Exc': excitation, 'Isec': i2100, 'Bdn': burdens, 'P1ap': eP1ap,
               'P1as': eP1as, 'P2ap': eP2ap, 'P2as': eP2as, 'P3ap': eP3ap, 'P3as': eP3as, 'P2asp': eP2asp, 'P1asp':eP1asp}
    print()
    print("Check caldict")

    print(Ta.check_caldict(caldict))
    print(Tb.calibrated)
    Tb.set_calibration(caldict)
    print(Tb.calibrated)
