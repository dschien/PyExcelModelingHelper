import importlib
from collections import defaultdict
from typing import Dict, List, Set

import numpy as np
import pandas as pd
from dateutil import relativedelta as rdelta
from openpyxl import load_workbook
import logging

__author__ = 'schien'

param_name_map = {'variable': 'name', 'scenario': 'source_scenarios_string', 'module': 'module_name',
                  'distribution': 'distribution_name', 'param 1': 'param_a', 'param 2': 'param_b', 'param 3': 'param_c',
                  'unit': '', 'CAGR': '', 'ref date': '', 'label': '', 'tags': '', 'comment': '', 'source': ''}

logging.basicConfig(level=logging.DEBUG)


class Parameter(object):
    """
    A single parameter
    """

    name: str
    unit: str
    comment: str
    source: str
    scenario: str

    "optional comma-separated list of tags"
    tags: str

    def __init__(self, name, value_generator=None, tags=None, source_scenarios_string: str = None, unit: str = None,
                 comment: str = None, source: str = None,
                 **kwargs):
        # The source definition of scenarios. A comma-separated list
        self.source = source
        self.comment = comment

        self.unit = unit
        self.source_scenarios_string = source_scenarios_string
        self.tags = tags
        self.name = name
        self.value_generator = value_generator

        self.scenario = None
        self.cache = None

    def __call__(self, *args, **kwargs):
        """
        Returns the same value everytime called.
        :param args:
        :param kwargs:
        :return:
        """
        if not self.cache:
            self.cache = self.value_generator.generate_values(*args, **kwargs)
        return self.cache


class DistributionFunctionGenerator(object):
    module: str
    distribution: str
    param_a: str
    param_b: str
    param_c: str

    def __init__(self, module_name=None, distribution_name=None, param_a: float = None,
                 param_b: float = None, param_c: float = None, size=None, **kwargs):
        self.size = size
        self.module_name = module_name
        self.distribution_name = distribution_name

        # prepare function arguments
        self.random_function_params = tuple([i for i in [param_a, param_b, param_c] if i])

    def generate_values(self, *args, **kwargs):
        """
        Generate values by sampling from a distribution. The size of the sample can be overriden with the 'size' kwarg.
        :param args:
        :param kwargs:
        :return:
        """
        f = self.instantiate_distribution_function(self.module_name, self.distribution_name)
        return f(*args if args else self.random_function_params, size=kwargs['size'] if 'size' in kwargs else self.size)

    @staticmethod
    def instantiate_distribution_function(module_name, distribution_name):
        module = importlib.import_module(module_name)
        func = getattr(module, distribution_name)
        return func


class ExponentialGrowthTimeSeriesGenerator(DistributionFunctionGenerator):
    cagr: str
    ref_date: str

    def __init__(self, cagr=None, times=None, size=None, index_names=None, ref_date=None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.cagr = cagr
        self.ref_date = ref_date
        self.times = times
        self.size = size
        iterables = [times, range(0, size)]
        self._multi_index = pd.MultiIndex.from_product(iterables, names=index_names)
        assert type(times.freq) == pd.tseries.offsets.MonthBegin, 'Time index must have monthly frequency'

    def __call__(self, *args, **kwargs):
        """
        Wraps the value in a dataframe
        :param name:
        :param kwargs:
        :return:
        """
        df = self.generate_values()
        df.set_index(self._multi_index, inplace=True)
        df.columns = [self.name]
        # @todo this is a hack to return a series with index as I don't know how to set an index and rename a series
        data_series = df.ix[:, 0]
        data_series._metadata = df._metadata
        return data_series

    def generate_values(self, *args, **kwargs):
        """
        Instantiate a random variable and apply annual growth factors.

        :return:
        """
        values = super().generate_values(*args, **kwargs, size=(len(self.times) * self.size,))
        alpha = self.cagr

        # @todo - fill to cover the entire time: define rules for filling first
        ref_date = self.ref_date if self.ref_date else self.times[0].to_pydatetime()
        assert ref_date >= self.times[0].to_pydatetime(), 'Ref date must be within variable time span.'
        assert ref_date <= self.times[-1].to_pydatetime(), 'Ref date must be within variable time span.'

        start_date = self.times[0].to_pydatetime()
        end_date = self.times[-1].to_pydatetime()

        a = growth_coefficients(start_date, end_date, ref_date, alpha, self.size)

        values *= a.ravel()

        df = pd.DataFrame(values)
        # df._metadata = [options]
        return df


def growth_coefficients(start_date, end_date, ref_date, alpha, samples):
    """
    Build a matrix of growth factors according to the CAGR formula  y'=y0 (1+a)^(t'-t0).
    The
    """
    if ref_date < start_date:
        raise ValueError("Ref data must be >= start date.")
    if ref_date > end_date:
        raise ValueError("Ref data must be >= start date.")
    if ref_date > start_date and alpha >= 1:
        raise ValueError("For a CAGR >= 1, ref date and start date must be the same.")
    # relative delta will be positive if ref date is >= start date (which it should with above assertions)
    delta_ar = rdelta.relativedelta(ref_date, start_date)
    ar = delta_ar.months + 12 * delta_ar.years
    delta_br = rdelta.relativedelta(end_date, ref_date)
    br = delta_br.months + 12 * delta_br.years

    # we place the ref point on the lower interval (delta_ar + 1) but let it start from 0
    # in turn we let the upper interval start from 1
    g = np.fromfunction(lambda i, j: np.power(1 - alpha, np.abs(i) / 12), (ar + 1, samples), dtype=float)
    h = np.fromfunction(lambda i, j: np.power(1 + alpha, np.abs(i + 1) / 12), (br, samples), dtype=float)
    g = np.flipud(g)
    # now join the two arrays
    return np.vstack((g, h))


class ParameterScenarioSet(object):
    """
    The set of all version of a parameter for all the scenarios.
    """
    default_scenario = 'default'

    "the name of the parameters in this set"
    parameter_name: str
    scenarios: Dict[str, Parameter]

    def __init__(self):
        self.scenarios = {}

    def add_scenario(self, parameter: 'Parameter', scenario_name: str = default_scenario):
        """
        Add a scenario for this parameter.

        :param scenario_name:
        :param parameter:
        :return:
        """
        self.scenarios[scenario_name] = parameter

    def __getitem__(self, item):
        return self.scenarios.__getitem__(item)

    def __setitem__(self, key, value):
        return self.scenarios.__setitem__(key, value)


class ParameterRepository(object):
    """
    Contains all known parameters.

    Internally, parameters are organised together with all the scenario variants in a single ParameterScenarioSet.

    """
    parameter_sets: Dict[str, ParameterScenarioSet]
    tags: Dict[str, Dict[str, Set[Parameter]]]

    def __init__(self):
        self.parameter_sets = defaultdict(ParameterScenarioSet)
        self.tags = defaultdict(lambda: defaultdict(set))

    def add_all(self, parameters: List[Parameter]):
        for p in parameters:
            self.add_parameter(p)

    def add_parameter(self, parameter: Parameter):
        """
        A parameter can have several scenarios. They are specified as a comma-separated list in a string.
        :param parameter:
        :return:
        """

        # try reading the scenarios from the function arg or from the parameter attribute
        scenario_string = parameter.source_scenarios_string
        if scenario_string:
            _scenarios = [i.strip() for i in scenario_string.split(',')]
            self.fill_missing_attributes_from_default_parameter(parameter)

        else:
            _scenarios = [ParameterScenarioSet.default_scenario]

        for scenario in _scenarios:
            parameter.scenario = scenario
            self.parameter_sets[parameter.name][scenario] = parameter

        # record all tags for this parameter
        if parameter.tags:
            _tags = [i.strip() for i in parameter.tags.split(',')]
            for tag in _tags:
                self.tags[tag][parameter.name].add(parameter)

    def fill_missing_attributes_from_default_parameter(self, param):
        """
        Empty fields in Parameter definitions in scenarios are populated with default values.

        E.g. in the example below, the source for the Power_TV variable in the 8K scenario would also be EnergyStar.

        | name     | scenario | val | tags   | source     |
        |----------|----------|-----|--------|------------|
        | Power_TV |          | 60  | UD, TV | EnergyStar |
        | Power_TV | 8K       | 85  | new_tag|            |

        **Note** tags must not differ. In the example above, the 8K scenario variable the tags value would be overwritten
        with the default value.

        :param param:
        :return:
        """
        default = self.parameter_sets[param.name][ParameterScenarioSet.default_scenario]
        for att_name, att_value in default.__dict__.items():
            if att_name in ['unit', 'label', 'comment', 'source', 'tags']:

                if att_name == 'tags' and default.tags != param.tags:
                    logging.warning(
                        f'For param {param.name} for scenarios {param.source_scenarios_string}, tags is different from default parameter tags. Overwriting with default values.')
                    setattr(param, att_name, att_value)

                if not getattr(param, att_name):
                    logging.info(
                        f'For param {param.name} for scenarios {param.source_scenarios_string}, populating attribute {att_name} with value {att_value} from default parameter.')

                    setattr(param, att_name, att_value)

    def __getitem__(self, item):
        return self.parameter_sets[item][ParameterScenarioSet.default_scenario]

    def get_parameter(self, param_name, scenario_name=None):
        return self.parameter_sets[param_name][
            scenario_name if scenario_name else ParameterScenarioSet.default_scenario]

    def find_by_tag(self, tag) -> Dict[str, Set[Parameter]]:
        """
        Get all registered dicts that are registered for a tag

        :param tag: str - single tag
        :return: a dict of {param name: set[Parameter]} that contains all ParameterScenarioSets for
        all parameter names with a given tag
        """
        return self.tags[tag]

    def exists(self, param):
        return param in self.parameter_sets.keys()


class ExcelParameterLoader(object):
    def __init__(self, filename, times=None, size=None, **kwargs):
        self.size = size
        if times is not None and size is None or times is None and size is not None:
            raise Exception('Both times and size arg must be set at the same time. Or none.')
        self.times = times
        self.filename = filename

    def load_parameter_definitions(self, sheet_name: str = None):
        """
        Load variable text from rows in excel file.
        :param sheet_name:
        :return:
        """
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

    def load_into_repo(self, repository: ParameterRepository = None, sheet_name: str = None):
        """
        Create a Repo from an excel file.
        :param repository: the repository to load into
        :param sheet_name:
        :return:
        """
        repository.add_all(self.load_parameters(sheet_name))

    def load_parameters(self, sheet_name):
        parameter_definitions = self.load_parameter_definitions(sheet_name=sheet_name)
        params = []
        for _def in parameter_definitions:

            # substitute names from the headers with the kwargs names in the Parameter and Distributions classes
            # e.g. 'variable' -> 'name', 'module' -> 'module_name', etc
            parameter_kwargs_def = {}
            for k, v in _def.items():
                if param_name_map[k]:
                    parameter_kwargs_def[param_name_map[k]] = v
                else:
                    parameter_kwargs_def[k] = v

            if self.times is not None:
                generator = ExponentialGrowthTimeSeriesGenerator(times=self.times,
                                                                 size=self.size, **parameter_kwargs_def)
            else:
                generator = DistributionFunctionGenerator(**parameter_kwargs_def)

            name_ = parameter_kwargs_def['name']
            del parameter_kwargs_def['name']
            p = Parameter(name_, **parameter_kwargs_def, value_generator=generator)
            params.append(p)
        return params
