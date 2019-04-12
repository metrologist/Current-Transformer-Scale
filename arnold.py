from __future__ import division
"""
Module to assist with managing readings from the Arnold test set. Start with
just assisting the usual spread sheet approach, but it might be possible to 
do more.
"""
import math
import GTC

class ARNOLD(object):
    """
    Analysis of the Arnold test set in MSLT.E.056.004 leads to a simple approach
    of applying single correction factors (one factor for all excitation levels)
    due to the auxiliary transformer. Corrections due to the slidewire and mutual
    inductor are set to zero, but uncertainty terms are added for resolution.
    """
    def __init__(self, source, output):
        """
        'source' is the spreadsheet with data and 'output' is the name of the
        spreadsheet the results are placed in.
        """
        self.source = source
        self.output = output
    
    def resolutionM(self, reading, multiplier):
        """
        equation from page 34 of MSLT.E.056.004
        multiplier returns it to 25 minute range (x1)
        reading is in crad, as is the returned value
        """
        #convert crad to min
        reading = reading/100/(2*math.pi)*360*60
        base_reading = reading/multiplier
        if abs(base_reading) >= 1.0:
            res = 0.04/math.log(25)*math.log(abs(base_reading)) + 0.01
        if abs(base_reading)< 1.0:
            res = 0.01
        res = res*multiplier
        #convert min to crad
        res = res/60/360*2*math.pi*100
        return res #in crad
        
    def corrnM(self, reading, multiplier):
        """
        simple uncertainty term for M scale correction
        calibrated on 25 minute range
        """
        res = 0.0017#crad
        return res*multiplier # in crad
        
    def resolutionR(self, reading, multiplier):
        """
        A simple 0.001% on the 0.5% range (x1)
        Not dependent on reading
        """
        res = 0.001
        return res*multiplier #in %
        
    def corrnR(self, reading, multiplier):
        """
        simple uncertainty term for R scale correction
        calibrated on 0.5% minute range
        """
        res = 0.0026/2.1 #% as typical uncertainty
        return res*multiplier # in %
        
    def auxE(self, reading):
        """
        uncertainty in the correction factor for auxiliary error
        """
        max_sd = 0.00007754
        return sd*reading
        
    def auxP(self, reading):
        """
        uncertainty in the correction factor for auxiliary phase
        """
        max_sd = 0.0003270
        return sd*reading

        
    

if __name__ == "__main__":
        print 'Testing that the module works as intended.'
        arnie = ARNOLD('a','b')
        #in GTC terms we add the corrections
        res_corrn_e = GTC.ureal(0, arnie.resolutionR(-0.682,1),5,'swire_resolution').u
        scale_corrn_e = GTC.ureal(0, arnie.corrnR(-0.682,1),5,'swire_correction').u
        res_corrn_p = GTC.ureal(0, arnie.resolutionM(0.032,0.1), 5,'mutual_resolution').u
        scale_corrn_p = GTC.ureal(0, arnie.corrnM(0.032,0.1), 5,'mutual_correction').u
        
        print res_corrn_e,'%'
        print scale_corrn_e,'%'
        print res_corrn_p,'crad'
        print scale_corrn_p,'crad'
        
"""
>pythonw -u "arnold.py"
Testing that the module works as intended.
0.001 %
0.00145444128109 crad
>Exit code: 0
"""