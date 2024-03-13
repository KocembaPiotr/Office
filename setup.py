from setuptools import setup, find_packages

VERSION = '0.0.82'
DESCRIPTION = 'Python package to manage office tools.'
LONG_DESCRIPTION = '''Python package which will simplify usage of main office tools.
                   '''

# Setting up
setup(
    name="pkoffice",
    version=VERSION,
    author="Piotr Kocemba",
    author_email="<KocembaPiotr@gmail.com>",
    description=DESCRIPTION,
    long_description=LONG_DESCRIPTION,
    packages=find_packages(),
    install_requires=[],
    keywords=['python', 'office', 'sql', 'excel', 'outlook'],
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Education",
        "Programming Language :: Python :: 2",
        "Programming Language :: Python :: 3",
        "Operating System :: MacOS :: MacOS X",
        "Operating System :: Microsoft :: Windows",
    ]
)