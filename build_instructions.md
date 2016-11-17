# Doc
[https://python-packaging-user-guide.readthedocs.org/en/latest/distributing.html#uploading-your-project-to-pypi]

# Build

`python setup.py sdist`

# Upload
It might be necessary to first delete old packages in 'dist' folder
`twine upload dist/*`

# Local install

## Build wheel
`python setup.py bdist_wheel -d dist`


## Pip install local wheel
`pip install dist/excel_modelling_helper-0.3.1-py3-none-any.whl --upgrade`

