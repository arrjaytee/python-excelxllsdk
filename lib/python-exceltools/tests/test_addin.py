import os
import os.path
import sys

def test_import():
    from exceltools import _addin

    print(dir(_addin))



def test_excel_regisiter():
    from xlwings import App
    from exceltools import _addin

    os.putenv("PATH", os.pathsep.join([sys.prefix, sys.base_prefix, os.getenv("PATH")]))

    with App(visible=True) as app:
        assert app.api.RegisterXLL(_addin.__file__)
