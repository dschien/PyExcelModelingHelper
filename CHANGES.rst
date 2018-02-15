15.5.2015   0.1.1   Renamed class to ParameterLoader
22.5.2015   0.1.2   Add sheet index as parameter to loader
11.1.2016   0.2.2   Added support to generate pandas dataframes, update to python 3
18.4.2016   0.2.7   Added new flag 'single_var' to freeze all variables except one to their mean value - use in sensitivity analysis.
19.8.2016   0.3.0   Upgrade to xarray 0.8.1
20.8.2016   0.3.1   Single var mean now analytical for choice, uniform, triangular and normal; trim white space from var names
4.07.2017   0.4.0   Rewrite with new API
4.07.2017   0.4.1   Added XLWings interface to read from Excel
14.09.2017   0.5.0   Delay sampling from data source until __call__ on Parameter.
16.2.2018   0.5.1   Fixed error in generation of random distributions with zero param values
