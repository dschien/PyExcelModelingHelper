# Doc
[https://python-packaging-user-guide.readthedocs.org/en/latest/distributing.html#uploading-your-project-to-pypi]

# Build

`python setup.py sdist`

# Upload
It might be necessary to first delete old packages in 'dist' folder
`twine upload dist/*`