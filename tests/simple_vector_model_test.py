from excel_helper.helper import ModelLoader

__author__ = 'schien'

import unittest
import numpy as np
import matplotlib.pyplot as plt



# class TestVectorModel(unittest.TestCase):
#     def test_model(self):
#         # plt.ion() # enables interactive mode
#         samples = 2
#
#         data = ModelLoader('test.xlsx', size=samples)
#         a = data['a']


if __name__ == '__main__':
    # init random generator
    # http://www.kevinsheppard.com/images/0/09/Python_introduction.pdf p. 225
    np.random.seed(123)

    # unittest.main()
    data = ModelLoader('test.xlsx', size=2)
    a = data['a']
    print a