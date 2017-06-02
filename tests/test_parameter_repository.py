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
        p = Parameter('p')
        r = Parameter('r')

        repo = ParameterRepository()
        repo.add_parameter(p, scenarios='test_scenario')
        repo.add_parameter(r, scenarios='test_scenario')

        _p = repo.get_parameter('p', 'test_scenario')
        assert _p.name == 'p'

    def test_multiple_scenarios(self):
        p = Parameter('test')

        repo = ParameterRepository()
        repo.add_parameter(p, scenarios='s1,s2')

        _p = repo.get_parameter('test', 's1')

        assert _p == p

    def test_get_by_tag(self):
        r = Parameter('r', tags='t2')

        repo = ParameterRepository()

        repo.add_parameter(r)
        param_sets_t2 = repo.find_by_tag('t2')
        assert len(param_sets_t2[r.name]) == 1

    def test_multiple_tags(self):
        p = Parameter('test', tags='t1,t2')
        r = Parameter('r', tags='t2')

        repo = ParameterRepository()
        repo.add_parameter(p)
        repo.add_parameter(r)

        param_sets_t1 = repo.find_by_tag('t1')
        assert p.name in param_sets_t1.keys()
        assert len(param_sets_t1[p.name]) == 1
        assert param_sets_t1[p.name].pop() == p.name

        param_sets_t2 = repo.find_by_tag('t2')
        assert len(param_sets_t2[p.name]) == 2

        assert [_set.pop() for _set in param_sets_t1[p.name]] == p.name


if __name__ == '__main__':
    unittest.main()
