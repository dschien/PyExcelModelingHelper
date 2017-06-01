import unittest

import pandas as pd
from scipy import stats
import math

from excel_helper.helper import Parameter, DistributionFunctionParameter, \
    ExponentialGrowthTimeSeriesDistributionFunctionParameter


class ParameterTestCase(unittest.TestCase):
    def test_distribution_generate_values(self):
        p = DistributionFunctionParameter('test', module_name='numpy.random', distribution_name='normal', param_a=0,
                                          param_b=.1, sample_size=32)

        a = p()
        assert abs(stats.shapiro(a)[0] - 0.9) < 0.1

    def test_ExponentialGrowthTimeSeriesDistributionFunctionParameter_generate_values(self):
        p = ExponentialGrowthTimeSeriesDistributionFunctionParameter('test', module_name='numpy.random',
                                                                     distribution_name='normal', param_a=0,
                                                                     param_b=.1,
                                                                     times=pd.date_range('2009-01-01', '2009-03-01',
                                                                                         freq='MS'),
                                                                     sample_size=5,
                                                                     cagr=0.1
                                                                     )

        a = p()
        print(a)
        # assert abs(stats.shapiro(a)[0] - 0.9) < 0.1


if __name__ == '__main__':
    unittest.main()
