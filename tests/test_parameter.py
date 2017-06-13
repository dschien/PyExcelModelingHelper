import unittest

import pandas as pd
from excel_helper import Parameter, DistributionFunctionGenerator, ExponentialGrowthTimeSeriesGenerator, \
    ParameterRepository, ExcelParameterLoader
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

    def test_get_mean_uniform(self):
        p = Parameter('a', value_generator=DistributionFunctionGenerator(module_name='numpy.random',
                                                                         distribution_name='normal', param_a=3,
                                                                         param_b=4, size=3, sample_mean_value=True))
        val = p()
        print(val)
        assert (val == 3).all()

    def test_get_mean_normal(self):
        assert False

    def test_get_mean_choice(self):
        assert False

    def test_get_mean_numerically(self):
        assert False


if __name__ == '__main__':
    unittest.main()
