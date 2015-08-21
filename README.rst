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

## Pip build local install wheel
`pip wheel -e dist/excel_modelling_helper-0.1.7-py2-none-any.whl`

## Pip install local wheel
`pip install wheelhouse/excel_modelling_helper-0.1.7-py2-none-any.whl`