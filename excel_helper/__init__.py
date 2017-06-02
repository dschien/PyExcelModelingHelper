from openpyxl import load_workbook

from excel_helper.helper import ParameterRepository, Parameter, ExponentialGrowthTimeSeriesGenerator, \
    DistributionFunctionGenerator

__author__ = 'schien'
import openpyxl

param_name_map = {'variable': 'name', 'scenario': 'source_scenarios_string', 'module': 'module_name',
                  'distribution': 'distribution_name', 'param 1': '', 'param 2': '', 'param 3': '', 'unit': '',
                  'CAGR': '', 'ref date': '', 'label': '', 'tags': '', 'comment': '', 'source': ''}


class ExcelParameterLoader(object):
    def __init__(self, filename, times=None, sample_size=None):
        self.sample_size = sample_size
        if times is not None and sample_size is None or times is None and sample_size is not None:
            raise Exception('Both times and sample_size arg must be set at the same time. Or none.')
        self.times = times
        self.filename = filename

    def load_parameter_definitions(self, sheet_name: str = None):
        definitions = []

        wb = load_workbook(filename=self.filename)
        _sheet_names = [sheet_name] if sheet_name else wb.sheetnames

        for _sheet_name in _sheet_names:
            sheet = wb.get_sheet_by_name(_sheet_name)
            rows = list(sheet.rows)
            header = [cell.value for cell in rows[0]]

            for row in rows[1:]:
                values = {}
                for key, cell in zip(header, row):
                    values[key] = cell.value
                definitions.append(values)

        return definitions

    def create_param_repo(self, sheet_name=None):
        repository = ParameterRepository()
        definitions = self.load_parameter_definitions(sheet_name=sheet_name)
        for _def in definitions:

            # substitute names from the headers with the kwargs names in the Parameter and Distributions classes
            _invert_def = {}
            for k, v in _def.items():
                if param_name_map[k]:
                    _invert_def[param_name_map[k]] = v
                else:
                    _invert_def[k] = v

            if self.times is not None:
                generator = ExponentialGrowthTimeSeriesGenerator(times=self.times,
                                                                 sample_size=self.sample_size, **_invert_def)
            else:
                generator = DistributionFunctionGenerator(**_invert_def)

            name_ = _invert_def['name']
            del _invert_def['name']
            p = Parameter(name_, **_invert_def, value_generator=generator)
            repository.add_parameter(p)
        return repository
