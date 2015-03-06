__author__ = 'schien'
import os

from setuptools import setup, find_packages


def read(*paths):
    """Build a file path from *paths* and return the contents."""
    with open(os.path.join(*paths), 'r') as f:
        return f.read()


setup(
    name='excel-modelling-helper',
    version='0.1.0',
    description='Use Excel to define your model parameters.',
    long_description=(read('README.rst') + '\n\n' +
                      read('CHANGES.rst')),
    url='http://github.com/dschien/PyExcelModellingHelper/',
    license='GPL, see LICENSE',
    author='Daniel Schien',
    author_email='dschien@gmail.com',
    py_modules=['excel-modelling-helper'],
    include_package_data=True,
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Science/Research',
        'Natural Language :: English',
        'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Topic :: Scientific/Engineering',
    ],
    packages=find_packages(exclude=['tests*']),
    install_requires=[
        'xlrd',
    ],
)