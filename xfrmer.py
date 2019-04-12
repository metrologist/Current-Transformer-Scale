from __future__ import division
from __future__ import print_function


class TRANSFORMER(object):
    """
    The TRANSFORMER class captures the essential structure of a transformer. It is specifically constructed for use
    with MSL's two-stage primary reference current transformers. These transformers have a main secondary and a second
    core auxiliary secondary. Primary windings are around both cores layered in groups with each group having a series
    or parallel connection. It should also be useable for a single core transformer with multiple secondary taps and
    fixed or window wound primary windings.

    There will likely be some experimentation with lists and dictionaries as to how best to describe a transfsormer.
    It should be possible to include calibration constants.
    """
    def __init__(self, primaries, secondaries, cores, type):
        """

        :param primaries: a list of sets of primary windings
        :param sedondaries: a list of sets of secondary windings
        :param cores: a list of cores (either 1 or 2 cores)
        :param type: either current or voltage
        """
        self.primaries = primaries
        self.secondaries = secondaries
        self.cores = cores
        assert type in ['voltage' , 'current'], "transformer type must be voltage or current"
        self.type = type

    def nominal_ratio(self, primary, secondary):
        if self.type == 'voltage':
            rat = primary/secondary
        elif self.type == 'current':
            rat = secondary/primary
        return rat

    def series(self, primary):
        return sum(primary)

    def parallel(self, primary):
        for i in range(1, len(primary)):
            assert primary[i] == primary[0], 'parallel windings must all have identical number of turns'
        return primary[0]