"""
Build python models using parameter values from Excel.

| variable | module | distribution      | param 1 | param 2 | param 3 | unit | label                      |
|----------|--------|-------------------|---------|---------|---------|------|----------------------------|
| c_ON     | numpy.random  | triangular | 0.25    | 0.75    | 1       | -    | number WDM terminals metro |

"""
from collections import defaultdict

import importlib
import itertools
import operator
import xlrd

__author__ = 'schien'

NAME = 'name'
DISTRIBUTION = 'distribution'
SCENARIO = 'scenario'
PARAM_A = 'param_a'
PARAM_B = 'param_b'
PARAM_C = 'param_c'
MODULE = 'module'
LABEL = 'label'
UNIT = 'unit'

HEADER_SEQ = [NAME, SCENARIO, MODULE, DISTRIBUTION, PARAM_A, PARAM_B, PARAM_C, UNIT, LABEL]

TABLE_STRUCT = {k: i for i, k in enumerate(HEADER_SEQ)}

DEFAULT_SCENARIO = 'def'


class ModelLoader(object):
    def __init__(self, file, size=1):
        wb = load_workbook(file)

        self.data = defaultdict(list)
        for key, group in itertools.groupby(wb, operator.itemgetter(1)):


            # scenario = row[1] if row[1] else DEFAULT_SCENARIO
            if not key:
                key = DEFAULT_SCENARIO
            self.data[key] = list(group)

            self.size = size
            self.cache = defaultdict(dict)

        self.scenario = DEFAULT_SCENARIO

    def select_scenario(self, scenario):
        self.scenario = scenario

    def unselect_scenario(self):
        self.scenario = DEFAULT_SCENARIO

    def get_row(self, name):
        scenario = self.scenario
        variables = [row[TABLE_STRUCT[NAME]] for row in self.data[scenario]]
        i = variables.index(name)
        return self.data[scenario][i]

    def get_val(self, name, args=None):
        """
        Apply function to arguments from excel table
         args: optional additonal args
        If no args are given, applies default size from constructor
        """
        scenario = self.scenario
        f, p = None, None
        if name in self.cache[scenario]:
            f, p = self.cache[scenario][name]
        else:
            row = self.get_row(name)
            f, p = build_distribution(row)
            self.cache[scenario][name] = (f, p)
        if args is not None:
            ret = f(*p, **args)
            assert ret.shape == (self.size,)
            return ret
        else:
            ret = f(*p, size=self.size)
            assert ret.shape == (self.size,)
            return ret

    def clear_cache(self):
        self.cache = defaultdict(dict)

    def get_label(self, name):
        try:
            row = self.get_row(name)
        except:
            return name
        return row[TABLE_STRUCT[LABEL]]

    def get_property(self, name, prop):
        try:
            row = self.get_row(name)
        except:
            return name
        return row[TABLE_STRUCT[prop]]

    def __getitem__(self, name):
        """
        Get the distribution for a item name from the table
        Then execute and return the result array
        """
        return self.get_val(name)


def build_distribution(row):
    module = importlib.import_module(row[TABLE_STRUCT[MODULE]])
    func = getattr(module, row[TABLE_STRUCT[DISTRIBUTION]])
    if row[TABLE_STRUCT[DISTRIBUTION]] == 'choice':

        cell = row[TABLE_STRUCT[PARAM_A]]
        if type(cell) in [float, int]:
            params = ([cell],)
        else:
            tokens = cell.split(',')

            params = [float(token.strip()) for token in tokens]
            params = (params, )
    elif row[TABLE_STRUCT[DISTRIBUTION]] == 'Distribution':
        func = func()
        params = tuple(row[TABLE_STRUCT[i]] for i in [PARAM_A, PARAM_B, PARAM_C] if row[TABLE_STRUCT[i]])
    else:
        params = tuple(row[TABLE_STRUCT[i]] for i in [PARAM_A, PARAM_B, PARAM_C] if row[TABLE_STRUCT[i]])
    return func, params


def load_workbook(file):
    """
    Build a list of tuples from the columns of the excel table.

    :param file:
    :return: a list of tuples
    [
        (col1,col2,...),
        (col1,col2,...),
    ]
    """
    wb = xlrd.open_workbook(file)

    sh = wb.sheet_by_index(0)

    rows_es = zip(*[sh.col_values(col_idx)[1:] for col_idx in range(0, len(HEADER_SEQ))])
    return rows_es