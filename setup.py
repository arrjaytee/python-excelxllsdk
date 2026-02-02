from setuptools import setup, Extension

setup(
    ext_modules=[
        Extension("pelib.test._test_pelib", ["pelib/test/_test_pelib.c"]),
    ],
    package_data={"": ["*.xls"], "ExcelXLLSDK.test": ["test_builtins/*.xls"]},
)
