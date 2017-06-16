__author__ = 'schien'
import os

import pypandoc
from setuptools import setup, find_packages


def read(*paths):
    """Build a file path from *paths* and return the contents."""
    with open(os.path.join(*paths), 'r') as f:
        return f.read()


long_description = pypandoc.convert('README.md', 'rst')

setup(
    name='excel-modelling-helper',
    version='0.4.0',
    description='Use Excel to define your model parameters.',
    long_description=(long_description + '\n\n' +
                      read('CHANGES.rst')),
    url='http://github.com/dschien/PyExcelModelingHelper/',
    license='GPL, see LICENSE',
    author='Daniel Schien',
    author_email='dschien@gmail.com',
    py_modules=['excel-modelling-helper'],
    download_url='https://github.com/dschien/PyExcelModelingHelper/releases/tag/0.2.0',
    include_package_data=True,
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Science/Research',
        'Natural Language :: English',
        'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Topic :: Scientific/Engineering',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.5',
    ],
    packages=find_packages(exclude=['tests*']),
    install_requires=[
        'xarray', 'xlrd', 'pandas', 'numpy', 'openpyxl', 'python-dateutil'
    ],
)
