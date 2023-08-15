del /f/s/q dist\*
py setup.py sdist bdist_wheel
twine upload dist/*