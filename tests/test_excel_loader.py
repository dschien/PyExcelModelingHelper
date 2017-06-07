import unittest

import pandas as pd

from excel_helper import ExcelParameterLoader, ParameterRepository


class ExcelParameterLoaderTestCase(unittest.TestCase):
    def test_load_by_sheetname(self):
        defs = ExcelParameterLoader(filename='./test_excelparameterloader.xlsx').load_parameter_definitions(
            sheet_name='Sheet1')
        for i, name in enumerate(['a', 'b', 'c']):
            assert defs[i]['variable'] == name

    def test_create_with_timeseries(self):
        loader = ExcelParameterLoader(filename='./test_excelparameterloader.xlsx',
                                      times=pd.date_range('2009-01-01', '2009-03-01', freq='MS'), size=10)
        repository = ParameterRepository()
        loader.load_into_repo(sheet_name='Sheet1', repository=repository)
        tag_param_dict = repository.find_by_tag('user')
        keys = tag_param_dict.keys()
        print(keys)
        assert 'a' in keys

        print(tag_param_dict['a'])


if __name__ == '__main__':
    unittest.main()
