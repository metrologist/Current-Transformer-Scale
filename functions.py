"""
This module holds all the functions used to model component behaviour in the
MODEL claas and also holds the non-linear least squares method used for fitting
these model functions to calibration data.
"""
import GTC as gtc #lower case for convenience only
import numpy as np #can differentiate, e.g. np.log from gtc.log
from scipy.optimize import leastsq

class FUNC(object):
    """
    A class to enable functions to return their number of coefficients, an
    identifying label and the actual function itself.  This helps with the
    automatic selection of functions and allows the formatting of files and
    tables for the correct number of coefficients.  It takes a method,
    *function*, together with the number of function coefficients, *no_coeffs*
    and a *label* to identify the method. These three arguments can then be
    accessed as needed
    """
    def __init__(self, function, no_coeffs, label):
        self.fn = function
        self.no = no_coeffs
        self.label = label

class PARAMETER(object):
    """
    A convenience for passing an initial value of a function coefficient to a
    fitting routine that then returns an improved estimate of the coefficient.
    It takes an *initial value* and a *name* and this value can be reset with
    the 'set' method.
    """
    def __init__(self, initialvalue, name):
        self.value = initialvalue
        self.name = name
    def set(self, value):
        self.value = value
    def __call__(self):
        return self.value

class MODEL(object):
    """
    A collection of functions that are used as the calibration equations for
    components in the metering installation. At initialisation, essential
    attributes of the fitting functions are gathered in 'FUNC' objects.  All the
    functions are constructed to return lists of values that are calculated for
    each item in the input list *X* or *r*.  Function coefficients are all
    entered as lists, *a*.
    """
    def __init__(self, name):
        self.name = name
        self.f_mean = FUNC(self.mean, 1, 'y = a0')
        self.f_line = FUNC(self.line, 2, 'y = a0 + a1*x')
        self.f_logline = FUNC(self.logline, 3, 'y = a0 + a1*x +a2*log(x)')
        self.f_smean = FUNC(self.smean, 1, 'z = a0')
        self.f_plane = FUNC(self.plane, 3, 'z = a0 + a1*x +a2*y')
        self.f_tan_s1 = FUNC(self.tan_s1, 3, 'z = a0 + a1*x + a2*tan(y)')
        self.f_tan_s2 = FUNC(self.tan_s2, 4, 'z= a0 + a1*x + (a2 + a3*log(x))*tan(y)')

#functions require X to be a list (1 or more values) and return a list
    def mean(self, X, a):
        """
        A single value, the mean of the input data.
        """
        #a0
        return [a[0] for x in X]

    def line(self, X, a):
        """
        A straight line, a0 + a1*x.
        """
        # a0 + a1*x
        return [a[0] + a[1]*x for x in X]

    def logline(self, X, a):
        """
        A function that has proved to be a good empirical fit to CT phase and
        error behaviour, a0 + a1*x + a2*log(x).
        """
        #a0 + a1*x + a2*log(x)
        #allow for x being an UncertainReal, else no uncertainty from log(x)
        if type(X[0]).__name__ == 'UncertainReal':
            package = gtc
        else:
            package = np
        return [a[0] + a[1]*x + a[2]*package.log(x) for x in X]

    def smean(self, r, a):
        """
        Bivariate version of 'mean'.
        """
        #assumes that r comes in as a list of tuples of (x,y)
        return [a[0] for x, y in r]

    def plane(self, r, a):
        """
        Bivariate version of 'line', a[0]+ a[1]*x + a[2]*y.
        """
        #assumes that r comes in as a list of tuples of (x,y)
        return [(a[0]+ a[1]*x + a[2]*y) for x, y in r]

    def tan_s1(self, r, a):
        """
        A function that has proved to be a good empirical fit to many meters.
        It assumes a linear current dependence and a fixed phase displacement
        error between the voltage and current channels of the meter,
        a[0]+ a[1]*x + a[2]*np.tan(y*np.pi/180).
        """
        #assumes that r comes in as a list of tuples of (x,y)
        return [(a[0]+ a[1]*x + a[2]*np.tan(y*np.pi/180)) for x, y in r]

    def tan_s2(self, r, a):
        """
        An extension of 'tan_s1' that allows the phase displacement to change
        as in a CT, a[0]+ a[1]*x + (a[2] + a[3]*np.log(x))*np.tan(y*np.pi/180).
        """
        #assumes that r comes in as a list of tuples of (x,y)
        return [(a[0]+ a[1]*x + (a[2] + a[3]*np.log(x))*np.tan(y*np.pi/180))
        for x, y in r]

    def influence1(self, w, w0, alpha):
        """
        A first order, simple linear temperature coefficient type function,
        alpha*(w - w0) . It assumes that *alpha* and (*w - w0*) are GTC ureals. Note
        that gtc.fn.mul2(a,b) is used to cover the possibility that both a and b
        have zero value, but a finite uncertainty.  If the value of both numbers
        is greater than twice its uncertainty, then normal multiplication is used.
        """
        trigger = 2.0 # could set the trigger as low as 1
        delta = w -w0
        if delta.u==0 or alpha.u==0:#some situations where unused influences given zero uncertainty
            return alpha*delta
        elif abs(delta.x)/delta.u < trigger or abs(alpha.x)/alpha.u < trigger:
            return gtc.fn.mul2(alpha, delta)
        else:
            return alpha*delta

    def fit(self, function, coefficients, x, data, u):
        """
        A list of initial *coefficients* of the *function* is passed to a
        Levenberg-Marquardt optimisation routine that returns  statistics
        of the least-squares fit and the fitted values of the coefficients.
        Input is in the form of a list of *data* values for independent
        varialble *x* and a list of their uncertainties (for weighting) *u*.
        The 'coefficents' are of PARAMETER class.
        """
        def fitfun(params):
            for i, p in enumerate(coefficients):
                p.set(params[i])
            return (data - function(x, params))/u
        if x is None:
            x = np.arange(data.shape[0])
        if u is None:
            u = np.ones(data.shape[0], "float")
        p = [param() for param in coefficients]
        return leastsq(fitfun, p, full_output=1, maxfev=0)

    def fit_ureals(self, coeffs, cov, dof, label):
        """
        Takes the coefficients, *coeffs*, and covariances, *cov*, together with
        degrees of freedom, *dof*, in the array output from *fit*  and returns a
        list of function coefficients as correlated GTC ureals.  The ureal is
        named with *label*.  If one of the uncertainties is zero, a set of unlinked
        ureals is created. This should only occur if a transformer has had all
        errors set to zero (e.g. no VT) or the input data has been fabricated to
        be a perfect fit.
        """
        uncert = []
        for i in range(len(coeffs)):
            uncert.append(np.sqrt(cov[i, i]))
        flag = 1 #assume no zero uncertainties
        for x in uncert:
            if x == 0.0:
                flag = 0
        if flag == 1:
            a = gtc.multiple_ureal(coeffs, uncert, dof, label)
            # and now set the correlations between all coefficients
            for i in range(len(coeffs)):
                for j in range(i):
                    gtc.set_correlation(cov[i, j]/np.sqrt(cov[i, i]*cov[j, j]), a[j], a[i])

        else: #have a zero uncertainty so create separate ureals
            print 'A fit has returned zero uncertainty for a function coefficient.  This should only occur if input data has been purposely set to a perfect fit.'
            a = []
            for i in range(len(coeffs)):
                a.append(gtc.ureal(coeffs[i], uncert[i], dof, label[i]))
        return a

    def fit_qual(self, variance, red_chisq, dof):
        """
        Checks the chi-square of the fit, *red_chisq*, and returns the
        uncertainty of an imposed type B component with a given value of zero.
        This uncertainty is also set to zero for a good fit, or a fit where the
        scatter is too high.  Both these cases are adequately accommodated by
        the usual least-squares process.

        If the chi-square has a probability of less than *prob* of being so small
        then it is identified as a "poor fit". The average *variance* is assigned
        to the type B component. The fit process in this case assumes that
        the calibration uncertainties were dominated by a type B component.

        If the chi-square has a probability of less than *prob* of being so
        large, then it is identified as a "bad fit", but the fit process assumes
        that the unexpected scatter is just an additional type A contribution.

        If the chi-square is seen as likely (1-*prob*), then the fit is
        identified as a "good fit". Note that scaling of covaricance by
        chi-square gives the correct confidence level.
        """
        prob = 0.1 #for now this is the trigger level determined to give acceptable long-run success
        fit_check = gtc.type_a.chisq_q(dof, red_chisq*dof)
        if fit_check > 1-prob:
            comment = 'Fit has lower uncertainty than expected from calibration data, so type B component added.'
            var_b0 = variance #
        elif fit_check < prob: # a bad fit to data!
            #cov = cov*red_chisq
            var_b0 = 0.0 #i.e. no offset uncertainty
            comment = 'Data scatter greater than expected, a poor fit'
        else:
            comment = 'Good fit.'
            var_b0 = 0.0#i.e. no offset uncertainty
        return [np.sqrt(var_b0), comment]

    def average_var(self, u, k):
        """
        Calculates the variance required as an input to 'fit_qual'.  It
        calculates the average of (*u*/*k*)**2, where *u* is a list of
        expanded uncertainties and *k* a list of coresponding coverage factors.
        An average standard deviation might be a better choice.
        """
        total = 0
        for i in range(len(k)):
            total = total + (u[i]/k[i])**2
        return total/len(k)

    def budget_by_label(self, target, labels):
        """
        A GTC calculated ureal *target* has its uncertainty reported against
        lists of labels in *labels*. Labels identify uncertainty inputs to
        target.  This facilitates reporting of uncertainties due to separate
        metering components.  There is no error checking so the onus is entirely
        on the programmer to choose relevant lists of labels for parameters.  A
        list of summed variance by label group is returned.
        """
        n = len(labels)# expect a list of lists
        total = []
        for i in range(n):
            total.append(0.0) # the uncertainty for each list will be summed
        for l, u in gtc.reporting.budget(target, trim = 0.0):
            for i in range(n):
                if l in labels[i]:
                    total[i] = total[i] + u**2
        return total




if __name__ == "__main__":
    print 'Testing components'
    model = MODEL('tran') # contains all fitting functions
    print model.line([1, 2], [3, 4])

    alpha = gtc.ureal(0.0e-6, 1.0e-6, 9)
    t = gtc.ureal(20.0, 0.5, 9)
    t0 = 20.0
    print model.influence1(t, t0, alpha).u
