import unittest

from excel_helper.helper import ParameterRepository, Parameter, ParameterScenarioSet


class ParameterRepositoryTestCase(unittest.TestCase):
    def test_add_parameter(self):
        p = Parameter('test')

        repo = ParameterRepository()
        repo.add_parameter(p)

    def test_add_scenario_parameter(self):
        """
        Test that missing properties are copied over from default parameter
        :return:
        """

    def test_get_parameter_getitem(self):
        p = Parameter('test')

        repo = ParameterRepository()
        repo.add_parameter(p)

        _p = repo['test']
        assert type(p) == Parameter
        assert _p == p

    def test_get_parameter_defaultscenario(self):
        p = Parameter('test')

        repo = ParameterRepository()
        repo.add_parameter(p)

        _p = repo.get_parameter('test')
        assert type(p) == Parameter
        assert _p == p

    def test_get_parameter_scenario(self):
        p = Parameter('test', scenario='test_scenario')
        r = Parameter('test')

        repo = ParameterRepository()
        repo.add_parameter(p)

        _p = repo.get_parameter('test', 'test_scenario')
        assert type(p) == Parameter
        assert _p == p


if __name__ == '__main__':
    unittest.main()
