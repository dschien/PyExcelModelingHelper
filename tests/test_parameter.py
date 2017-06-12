import unittest

import pandas as pd
from excel_helper import Parameter, DistributionFunctionGenerator, ExponentialGrowthTimeSeriesGenerator
from scipy import stats


class ParameterTestCase(unittest.TestCase):
    def test_distribution_generate_values(self):
        p = Parameter('test', value_generator=DistributionFunctionGenerator(module_name='numpy.random',
                                                                            distribution_name='normal', param_a=0,
                                                                            param_b=.1, size=32))

        a = p()
        assert abs(stats.shapiro(a)[0] - 0.9) < 0.1

    def test_ExponentialGrowthTimeSeriesDistributionFunctionParameter_generate_values(self):
        p = Parameter('test', value_generator=ExponentialGrowthTimeSeriesGenerator(module_name='numpy.random',
                                                                                   distribution_name='normal',
                                                                                   param_a=0,
                                                                                   param_b=.1,
                                                                                   times=pd.date_range('2009-01-01',
                                                                                                       '2009-03-01',
                                                                                                       freq='MS'),
                                                                                   size=5,
                                                                                   cagr=0.1
                                                                                   ))

        a = p()
        print(a)
        # assert abs(stats.shapiro(a)[0] - 0.9) < 0.1

    def test_normal_zero_variance(self):
        p = Parameter('a', value_generator=DistributionFunctionGenerator(module_name='numpy.random',
                                                                         distribution_name='normal', param_a=0,
                                                                         param_b=0, size=3))
        q = Parameter('b', value_generator=DistributionFunctionGenerator(module_name='numpy.random',
                                                                         distribution_name='normal', param_a=0,
                                                                         param_b=0, size=3))

        a = p() * q()
        assert abs(stats.shapiro(a)[0] - 0.9) < 0.1


if __name__ == '__main__':
    unittest.main()
