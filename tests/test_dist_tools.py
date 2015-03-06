import numpy as np
from excel_helper.helper import ModelLoader, build_distribution

__author__ = 'schien'

import unittest


class TestExcelTool(unittest.TestCase):
    # def setUp(self):
    # self.seq = range(10)

    def test_get_row(self):
        data = ModelLoader('test.xlsx')

        row = data.get_row('a')
        print row
        f, p = build_distribution(row)
        print p
        print f(*p)

    def test_choice(self):
        data = ModelLoader('test.xlsx', size=1)
        res = data['a'][0]
        assert res == 1.

    def test_cache(self):
        data = ModelLoader('test.xlsx', size=1)
        data['a'][0]
        res = data['a'][0]
        assert res == 1.

    def test_set_scenarios(self):
        data = ModelLoader('test.xlsx', size=1)
        res = data['a'][0]

        assert res == 1.

        data.select_scenario('s1')
        res = data['a'][0]

        assert res == 2.


    def test_unset_scenarios(self):
        data = ModelLoader('test.xlsx', size=1)
        res = data['a'][0]

        assert res == 1.

        data.select_scenario('s1')
        res = data['a'][0]

        assert res == 2.

        data.unselect_scenario()
        res = data['a'][0]

        assert res == 1.

    def test_cache_with_scenarios(self):
        data = ModelLoader('test.xlsx', size=1)
        data['a'][0]
        res = data['a'][0]

        assert res == 1.

        data.select_scenario('s1')
        data['a'][0]
        res = data['a'][0]

        assert res == 2.

    def test_arrays(self):
        data = ModelLoader('test.xlsx', size=2)
        res = data['a']

        assert np.array_equal(res, [1., 1.])

    def test_uniform(self):
        data = ModelLoader('test.xlsx', size=1)
        res = data['b'][0]

        assert (res < 4.) & (res > 2.)

    def test_triangular(self):
        data = ModelLoader('test.xlsx', size=1)
        res = data['c'][0]

        assert (res < 10.) & (res > 3.)


        # def test_module_choice(self):
        # data = ModelLoader('../data/metro_core_network_model_params.xlsx')
        #     res = data['R']
        #     print res


        # def test_module_args(self):
        #     data = ModelLoader('../data/metro_core_network_model_params.xlsx')
        #     eta_M_R = data.get_val('eta_M_R', {'size': 5})  # - $\eta_{M_R}$ metro router overcapacity/ utilisation
        #     print eta_M_R


if __name__ == '__main__':
    unittest.main()