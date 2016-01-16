from datetime import date, datetime
from dateutil import relativedelta
import numpy as np
from excel_helper.helper import ParameterLoader, build_distribution, DataSeriesLoader, growth_coefficients, \
    MultiSourceLoader, MCDataset
import pandas as pd

__author__ = 'schien'

import unittest


class TestExcelTool(unittest.TestCase):
    # def setUp(self):
    # self.seq = range(10)

    def test_open_by_name(self):
        data = ParameterLoader.from_excel('test.xlsx', size=1, sheet_name='Sheet1')
        res = data['c'][0]

        assert (res < 10.) & (res > 3.)

    def test_open_index_zero(self):
        data = ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=0)
        res = data['c'][0]

        assert (res < 10.) & (res > 3.)

    def test_get_row(self):
        data = ParameterLoader.from_excel('test.xlsx', sheet_index=0)

        row = data.get_row('a')
        # print row
        f, p, _ = build_distribution(row)
        print(p)
        print(f(*p))

    def test_cache(self):
        data = ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=0)
        data['a'][0]
        res = data['a'][0]
        assert res == 1.

    def test_set_scenarios(self):
        data = ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=0)
        res = data['a'][0]

        assert res == 1.

        data.select_scenario('s1')
        res = data['a'][0]

        assert res == 2.

    def test_unset_scenarios(self):
        data = ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=0)
        res = data['a'][0]

        assert res == 1.

        data.select_scenario('s1')
        res = data['a'][0]

        assert res == 2.

        data.unselect_scenario()
        res = data['a'][0]

        assert res == 1.

    def test_cache_with_scenarios(self):
        data = ParameterLoader.from_excel('test.xlsx', size=100000, sheet_index=0)
        a = data['a'][0]
        b = data['b'][0]

        c = a + b
        res = data['a'][0]

        assert res == 1.

        data.select_scenario('s1')
        data['a'][0]
        res = data['a'][0]

        assert res == 2.


class TestDistributionFunctions(unittest.TestCase):
    def test_choice(self):
        data = ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=0)
        res = data['a'][0]
        assert res == 1.

    def test_multiple_choice(self):
        data = ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=0)
        res = data['multiple choice'][0]
        assert res in [1., 2., 3.]

    def test_arrays(self):
        data = ParameterLoader.from_excel('test.xlsx', size=2, sheet_index=0)
        res = data['a']
        # print res
        assert np.array_equal(res, [1., 1.])

    def test_uniform(self):
        data = ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=0)
        res = data['b'][0]

        assert (res < 4.) & (res > 2.)

    def test_triangular(self):
        data = ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=0)
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


class TestDataFrameLoader(unittest.TestCase):
    def test_get_dataframe(self):
        # the time axis of our dataset
        times = pd.date_range('2009-01-01', '2009-04-01', freq='MS')
        # the sample axis our dataset
        samples = 3

        dfl = DataSeriesLoader.from_excel('test.xlsx', times, size=samples, sheet_index=0)
        res = dfl['a']

        # assert np.array_equal(res, [1., 1.])

    def test_metadata(self):
        # the time axis of our dataset
        times = pd.date_range('2009-01-01', '2009-04-01', freq='MS')
        # the sample axis our dataset
        samples = 3

        dfl = DataSeriesLoader.from_excel('test.xlsx', times, size=samples, sheet_index=0)
        res = dfl['a']

        print(res._metadata)

    def test_from_dataframe(self):
        times = pd.date_range('2009-01-01', '2009-04-01', freq='MS')

        xls = pd.ExcelFile('test.xlsx')
        df = xls.parse('Sheet1')
        ldr = DataSeriesLoader.from_dataframe(df, times, size=3)
        res = ldr['a']
        print(res)


class TestCAGRCalculation(unittest.TestCase):
    def test_one_month(self):
        """
        If start and end are identical, we expect an array of one row of ones of sample size

        :return:
        """
        samples = 3
        alpha = 1  # 100 percent p.a.
        ref_date = date(2009, 1, 1)
        start_date = date(2009, 1, 1)
        end_date = date(2009, 1, 1)

        a = growth_coefficients(start_date, end_date, ref_date, alpha, samples)
        assert np.all(a == np.ones((samples, 1)))

    def test_two_months(self):
        """
        If start and end are one month apart, we expect an array of one row of ones of sample size for the ref month
        and one row with CAGR applied

        :return:
        """
        samples = 3
        alpha = 1  # 100 percent p.a.
        ref_date = date(2009, 1, 1)
        start_date = date(2009, 1, 1)
        end_date = date(2010, 1, 1)

        a = growth_coefficients(start_date, end_date, ref_date, alpha, samples)
        assert np.all(a[0] == np.ones((samples, 1)))
        assert np.all(a[1] == np.ones((samples, 1)) * pow(1 + alpha, 1. / 12))

    def test_negative_growth(self):
        """
        If start and end are one month apart, we expect an array of one row of ones of sample size for the ref month
        and one row with CAGR applied

        :return:
        """
        samples = 3
        alpha = -0.1  # 100 percent p.a.
        ref_date = date(2009, 1, 1)
        start_date = date(2009, 1, 1)
        end_date = date(2010, 1, 1)
        a = growth_coefficients(start_date, end_date, ref_date, alpha, samples)
        # print a
        assert np.all(a[0] == np.ones((samples, 1)))
        assert np.all(a[-1] == np.ones((samples, 1)) * 1 + alpha)

    def test_refdate_after_start(self):
        """
        If the ref date is greater than the start, we expect an array of one row of ones of sample size for the ref month
        and rows with (1-CAGR)^t applied

        :return:
        """
        samples = 3
        alpha = 0.1  # 10 percent p.a.
        ref_date = date(2009, 2, 1)
        start_date = date(2009, 1, 1)
        end_date = date(2009, 2, 1)

        a = growth_coefficients(start_date, end_date, ref_date, alpha, samples)
        assert a.shape == (2, samples)
        # first row has negative coefficients
        assert np.all(a[0] == np.ones((samples, 1)) * pow(1 - alpha, 1. / 12))
        # second row has ref values
        assert np.all(a[-1] == np.ones((samples, 1)))

    def test_refdate_between_start_and_end(self):
        """
        If the ref date is greater than the start, we expect an array of one row of ones of sample size for the ref month
        and rows with (1-CAGR)^t applied

        :return:
        """
        samples = 3
        alpha = 0.1  # 10 percent p.a.
        ref_date = date(2009, 3, 1)
        start_date = date(2009, 1, 1)
        end_date = date(2009, 6, 1)

        delta = relativedelta.relativedelta(end_date, start_date)
        total_months = delta.months + delta.years * 12 + 1

        ref_delta = relativedelta.relativedelta(ref_date, start_date)
        ref_row_idx = ref_delta.months + ref_delta.years * 12

        a = growth_coefficients(start_date, end_date, ref_date, alpha, samples)
        # print a

        assert a.shape == (total_months, samples)
        # the ref_row_idx has all ones
        assert np.all(a[ref_row_idx] == np.ones((samples, 1)))
        # the row before ref_row_idx has negative coefficients
        assert np.all(a[ref_row_idx - 1] == np.ones((samples, 1)) * pow(1 - alpha, 1. / 12))
        # the row after ref_row_idx has positive coefficients
        assert np.all(a[ref_row_idx + 1] == np.ones((samples, 1)) * pow(1 + alpha, 1. / 12))

        # the last row has positive coefficients
        assert np.all(a[-1] == np.ones((samples, 1)) * pow(1 + alpha, float(total_months - 1 - ref_row_idx) / 12))


class TestDataFrameWithCAGRCalculation(unittest.TestCase):
    def test_simple_CAGR(self):
        """
        Basic test case, applying CAGR to a Pandas Dataframe.

        :return:
        """
        # the time axis of our dataset
        times = pd.date_range('2009-01-01', '2009-04-01', freq='MS')
        # the sample axis our dataset
        samples = 2

        dfl = DataSeriesLoader.from_excel('test.xlsx', times, size=samples, sheet_index=0)
        res = dfl['a']

        assert res.loc[[datetime(2009, 1, 1)]][0] == 1
        assert np.abs(res.loc[[datetime(2009, 4, 1)]][0] - pow(1.1, 3. / 12)) < 0.00001

    def test_simple_CAGR_from_pandas(self):
        times = pd.date_range('2009-01-01', '2009-04-01', freq='MS')

        xls = pd.ExcelFile('test.xlsx')
        df = xls.parse('Sheet1')
        ldr = DataSeriesLoader.from_dataframe(df, times, size=2)
        res = ldr['a']

        assert res.loc[[datetime(2009, 1, 1)]][0] == 1
        assert np.abs(res.loc[[datetime(2009, 4, 1)]][0] - pow(1.1, 3. / 12)) < 0.00001


class TestMultiSourceLoader(unittest.TestCase):
    def test_simple(self):
        data = MultiSourceLoader()

        data.add_source(ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=0))
        data.add_source(ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=1))
        a = data['a'][0]
        assert a == 1.

        z = data['z'][0]
        assert z == 1.

    def test_mixed_type_multisource(self):
        data = MultiSourceLoader()

        times = pd.date_range('2009-01-01', '2009-04-01', freq='MS')
        # the sample axis our dataset
        samples = 2

        dfl = DataSeriesLoader.from_excel('test.xlsx', times, size=samples, sheet_index=0)

        data.add_source(ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=1))
        res = dfl['a']

        assert res.loc[[datetime(2009, 1, 1)]][0] == 1
        assert np.abs(res.loc[[datetime(2009, 4, 1)]][0] - pow(1.1, 3. / 12)) < 0.00001

        z = data['z'][0]
        assert z == 1.

    def test_scenario(self):
        data = MultiSourceLoader()
        data.add_source(ParameterLoader.from_excel('test.xlsx', size=1, sheet_index=0))

        res = data['a'][0]

        assert res == 1.

        data.set_scenario('s1')
        res = data['a'][0]

        assert res == 2.

        data.reset_scenario()
        res = data['a'][0]

        assert res == 1.


class TestMCDataset(unittest.TestCase):
    def test_simple(self):
        data = MCDataset()
        times = pd.date_range('2009-01-01', '2009-04-01', freq='MS')
        # the sample axis our dataset
        samples = 2
        data.add_source(DataSeriesLoader.from_excel('test.xlsx', times, size=samples, sheet_index=0))
        # data.prepare('a')
        res = data['a']
        print(res)
        res = data.b
        print(res)
        # assert res.loc[[datetime(2009, 1, 1)]][0] == 1
        # assert np.abs(res.loc[[datetime(2009, 4, 1)]][0] - pow(1.1, 3. / 12)) < 0.00001


if __name__ == '__main__':
    unittest.main()
