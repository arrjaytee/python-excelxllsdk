import unittest
from unittest import skipIf
from ExcelXLLSDK.xltypes import  _malloc_errcheck

class xltypesTestCase(unittest.TestCase):
    def test_malloc_weak(self):
        """check that malloc checking is ok"""        
        with self.assertRaises(MemoryError):
            _malloc_errcheck(None, None, None)            
                       
    
