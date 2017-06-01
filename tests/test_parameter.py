import unittest
from scipy import stats
import math

from excel_helper.helper import Parameter, DistributionFunctionParameter


class ParameterTestCase(unittest.TestCase):
    def test_generate_values(self):
        p = DistributionFunctionParameter('test', module='numpy.random', distribution_name='normal', param_a=0,
                                          param_b=.1)

        a = p(size=32)
        assert abs(stats.shapiro(a)[0] - 0.9) < 0.1


if __name__ == '__main__':
    unittest.main()
