def test_import():
    from exceltools import _addin

    print(dir(_addin))


def test_excel_regisiter():
    from xlwings import App
    from exceltools import _addin

    with App(visible=True) as app:
        app.api.RegisterXLL(_addin.__file__)
