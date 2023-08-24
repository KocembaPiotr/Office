import os

os.system(r"del /f/s/q dist\*")
os.system(r"py setup.py sdist bdist_wheel")
os.system(r"twine upload dist/*")