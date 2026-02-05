"""
install exceltools, a package for deploying python excel extensions
"""

from setuptools import setup, Extension, find_packages

setup(
    ext_modules=[
        Extension("exceltools._addin", ["exceltools/_addin.c"]),
    ],
    packages=find_packages(),
)
