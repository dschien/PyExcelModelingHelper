import unittest

import pandas as pd

from excel_helper import ExcelParameterLoader, ParameterRepository


class ExcelParameterLoaderTestCase(unittest.TestCase):
    def test_load_by_sheetname(self):
        defs = ExcelParameterLoader(filename='./test_excelparameterloader.xlsx').load_parameter_definitions(
            sheet_name='Sheet1')
        for i, name in enumerate(['a', 'b', 'c']):
            assert defs[i]['variable'] == name

    def test_column_order(self):
        repository = ParameterRepository()
        ExcelParameterLoader(filename='./test_excelparameterloader.xlsx').load_into_repo(sheet_name='shuffle_col_order',
                                                                                         repository=repository)

        p = repository.get_parameter('z')
        assert p.name == 'z'
        assert p.tags == 'x'

    def test_choice_single_param(self):
        repository = ParameterRepository()
        ExcelParameterLoader(filename='./test_excelparameterloader.xlsx').load_into_repo(sheet_name='Sheet1',
                                                                                         repository=repository)
        p = repository.get_parameter('choice_var')
        assert p() == .9

    def test_choice_two_params(self):
        repository = ParameterRepository()
        ExcelParameterLoader(filename='./test_excelparameterloader.xlsx').load_into_repo(sheet_name='Sheet1',
                                                                                         repository=repository)
        p = repository.get_parameter('a')
        assert p() in [1, 2]

    def test_multiple_choice(self):
        repository = ParameterRepository()
        ExcelParameterLoader(filename='./test_excelparameterloader.xlsx').load_into_repo(sheet_name='Sheet1',
                                                                                         repository=repository)
        p = repository.get_parameter('multiple_choice')
        assert p() in [1, 2, 3]

    def test_choice_time(self):
        repository = ParameterRepository()
        ExcelParameterLoader(filename='./test_excelparameterloader.xlsx',
                             times=pd.date_range('2009-01-01', '2015-05-01', freq='MS'), size=10
                             ).load_into_repo(sheet_name='Sheet1', repository=repository)

        p = repository.get_parameter('choice_var')
        val = p()
        # print(val)
        assert (val == .9).all()

    def test_choice_two_params_with_time(self):
        loader = ExcelParameterLoader(filename='./test_excelparameterloader.xlsx',
                                      times=pd.date_range('2009-01-01', '2009-03-01', freq='MS'), size=10)
        repository = ParameterRepository()
        loader.load_into_repo(sheet_name='Sheet1', repository=repository)
        tag_param_dict = repository.find_by_tag('user')
        keys = tag_param_dict.keys()
        print(keys)
        assert 'a' in keys
        repository['a']()

        print(tag_param_dict['a'])

    def test_uniform(self):
        repository = ParameterRepository()
        ExcelParameterLoader(filename='./test_excelparameterloader.xlsx').load_into_repo(sheet_name='Sheet1',
                                                                                         repository=repository)
        p = repository.get_parameter('b')
        val = p()
        assert (val >= 2) & (val <= 4)

    def test_uniform_time(self):
        repository = ParameterRepository()
        ExcelParameterLoader(filename='./test_excelparameterloader.xlsx',
                             times=pd.date_range('2009-01-01', '2015-05-01', freq='MS'), size=10
                             ).load_into_repo(sheet_name='Sheet1', repository=repository)
        p = repository.get_parameter('b')
        val = p()
        print(val)
        print(type(val))

        assert (val >= 2).all() & (val <= 4).all()

    def test_triagular(self):
        repository = ParameterRepository()
        ExcelParameterLoader(filename='./test_excelparameterloader.xlsx').load_into_repo(sheet_name='Sheet1',
                                                                                         repository=repository)
        p = repository.get_parameter('c')

        res = p()

        assert (res < 10.) & (res > 3.)

    def test_triagular_time(self):
        repository = ParameterRepository()
        ExcelParameterLoader(filename='./test_excelparameterloader.xlsx',
                             times=pd.date_range('2009-01-01', '2015-05-01', freq='MS'), size=10
                             ).load_into_repo(sheet_name='Sheet1', repository=repository)

        p = repository.get_parameter('c')

        res = p()

        assert (res < 10.).all() & (res > 3.).all()

    def test_normal(self):
        pass


if __name__ == '__main__':
    unittest.main()
