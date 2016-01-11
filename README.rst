# Example
Given an excel file with rows similar to the below

| variable        | scenario | module                                      | distribution | param 1          | param 2 | param 3 | unit | start date | end date   | CAGR | ref date   | label      | comment | source |
|-----------------|----------|---------------------------------------------|--------------|------------------|---------|---------|------|------------|------------|------|------------|------------|---------|--------|
| a               |          | numpy.random                                | choice       | 1                |         |         | kg   | 01/01/2009 | 01/04/2009 | 0.10 | 01/01/2009 | test var 1 |         |        |
| b               |          | numpy.random                                | uniform      | 2                | 4       |         | -    |            |            |      |            | label      |         |        |
| c               |          | numpy.random                                | triangular   | 3                | 6       | 10      | -    |            |            |      |            | label      |         |        |
| d               |          | bottom_up_comparision.sampling_core_routers | Distribution | core_routers.csv |         |         | J/Gb |            |            |      |            | label      |         |        |
|                 |          |                                             |              |                  |         |         |      |            |            |      |            |            |         |        |
| a               | s1       | numpy.random                                | choice       | 2                |         |         |      |            |            |      |            | test var 1 |         |        |
| multiple choice |          | numpy.random                                | choice       | 1,2,3            |         |         | kg   | 01/01/2007 | 01/01/2009 |      |            | test var 1 |         |        |


You can run python/ numpy code that references these variables and generates random distributions.

For example, the following will initialise a variable `c` with a vector of size 2 with random values
 from a triangular distribution.

```
np.random.seed(123)

data = ParameterLoader.from_excel('test.xlsx', size=2, sheet_index=0)
c = data['c']
>>> [ 7.08471918  5.45131111]
```

Other types of distributions include `choice` and `normal`. However you can specify any distribution from
numpy that takes up to three parameters to init.

You can also specify a .csv file with samples and an empiricial distribution function is generated
and variable values will be sampled from that.