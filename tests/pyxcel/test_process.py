
from pyxcel.process import _get_executable, _get_installed_versions, ExcelPopen
from ExcelXLLSDK.XLCALL import xlver

import unittest

_versions = [ '11.0', '12.0', '14.0' ]
_versions = ['14.0']

@unittest.skipIf(xlver, "no need to check excel subprocess in excel")
class ExcelProcessTestCase(unittest.TestCase):
    def test_get_executable(self):
        """check that all required versions of excel are installed in the standard locations"""
        for ver in _versions:
            self.assertTrue(_get_executable(ver).endswith(u'Microsoft Office\\Office14\\EXCEL.EXE'))
        
    def test_installed_versions(self):
        """see what we have installed"""        
        self.assertEqual(list(_get_installed_versions()), _versions)

    def test_ExcelPopen(self):
        """check that excel popen gives us some unique processes"""
        def test(version):        
            # launch three processes
            ps = [ ExcelPopen(version=version) for _ in xrange(0,2) ]
        
            # check that they are all distinct instances, only way to tell
            # from the process itself is the pid
            Hwnds = [ p.Application.Hwnd for p in ps ]
            self.assertEqual(len(set(Hwnds)), len(Hwnds))

            # shut them all down nicely, and check they have gone
            for p in ps:
                p.quit()
                self.assertNotEqual(p.returncode, None)          

        for ver in _get_installed_versions():
            test(ver)


