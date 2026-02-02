"""
install exceltools, a package for deploying python excel extensions
"""

from setuptools import setup, Extension, find_packages

setup(
    ext_modules=[
        Extension("exceltools._addin", ["exceltools/_addin.c"]),
    ],
    packages=find_packages(),
    package_data={"": ["*.xll"]},
    entry_points={
        "console_scripts": [
            "excel_entry_points = exceltools.entrypoints:write_entry_points"
        ]
    },
)
