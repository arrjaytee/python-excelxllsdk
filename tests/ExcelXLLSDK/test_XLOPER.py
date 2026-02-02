#pylint: disable=C0111
from __future__ import absolute_import
import unittest

from ExcelXLLSDK.xltypes import XLOPER4, XLOPER12
from ExcelXLLSDK.gen.xlerr import xlerrNA, xlerrRef

from datetime import datetime

def _make(XLOPER):
    class _base(unittest.TestCase):
        def test_str(self):
            self.assertEqual(str(XLOPER(())), 'Missing')
            self.assertEqual(str(XLOPER(None)), 'Nil')
            self.assertEqual(str(XLOPER(float('NaN'))), '#N/A')
            self.assertEqual(str(XLOPER(123)), '123')
            self.assertEqual(str(XLOPER(123L)), '123')
            self.assertEqual(str(XLOPER(123.123)), '123.123')
            self.assertEqual(str(XLOPER("fish")), 'fish')
            self.assertEqual(str(XLOPER(True)), "TRUE")
            self.assertEqual(str(XLOPER(False)), "FALSE")
            self.assertEqual(str(XLOPER([True, 1, 1.1, 'fish'])), '{ TRUE; 1; 1.1; "fish" }')
            self.assertEqual(str(XLOPER(ZeroDivisionError())), '#DIV/0!')
            self.assertEqual(str(XLOPER(AssertionError())), '#VALUE!')
            self.assertEqual(str(XLOPER(NameError())), '#NAME?')
            self.assertEqual(str(XLOPER(TypeError())), '#VALUE!')
            self.assertEqual(str(XLOPER(StandardError())), '#NUM!')
            self.assertEqual(str(XLOPER(Exception())), '#NULL!')
            self.assertEqual(str(XLOPER([[1, 2], [3, 4]])), "{ 1, 2; 3, 4 }")
            self.assertEqual(str(XLOPER([[1, None], [3, 4]])), "{ 1, Nil; 3, 4 }")
            self.assertEqual(str(XLOPER([[1, float('NaN')], [3, 4]])), "{ 1, #N/A; 3, 4 }")
            

            xlo = XLOPER(RuntimeError())
            xlo.val.err = 0xFFFF
            self.assertEqual(str(xlo), '#UNKNOWN!')
            xlo.val.err = xlerrNA
            self.assertEqual(str(xlo), '#N/A')
            xlo.val.err = xlerrRef
            self.assertEqual(str(xlo), '#REF!')

        def test_datetime(self):
            dt = datetime.now().replace(microsecond=0)
            self.assertEqual(XLOPER(dt).datetime, dt)

        def test_repr(self):
            self.assertEqual(repr(XLOPER("fish")), '"fish"')

        def test_nonzero(self):
            self.assertTrue(XLOPER(True))
            self.assertTrue(not XLOPER(False))
            self.assertTrue(XLOPER([[1]]))
            self.assertTrue(not XLOPER([]))
            self.assertTrue(not XLOPER([()]))

        def test_str_length(self, max=0xFF):
            longstr = 'x' * max
            self.assertEqual(str(XLOPER(longstr)), longstr)
            with self.assertRaises(ValueError):
                XLOPER(longstr + 'x')
       
        def test_Missing(self):
            self.assertEqual(XLOPER(()).value, None)

        def test_None(self):
            self.assertEqual(XLOPER(None).value, None)

        def test_int(self):
            self.assertEqual(int(XLOPER(123)), 123)
            self.assertEqual("%d" % XLOPER(123), "123")

        def test_long(self):
            self.assertEqual(int(XLOPER(123L)), 123)
            self.assertEqual("%d" % XLOPER(123L), "123")
            self.assertRaises(TypeError, XLOPER, 123123123123123L)
            self.assertRaises(TypeError, XLOPER, -123123123123123L)

        def test_float(self):
            self.assertEqual(float(XLOPER(123.123)), 123.123)
            self.assertEqual("%f" % XLOPER(123.123), "123.123000")

        def test_string(self):
            self.assertEqual(str(XLOPER('')), '')
            self.assertEqual(str(XLOPER('fish')), 'fish')
            self.assertEqual("%s" % XLOPER('fish'), 'fish')

        def test_Multi(self):
            multi = XLOPER([True, 1, 1.1, "fish"])

            self.assertEqual(multi[0, 0], True)
            self.assertEqual(multi[1, 0], 1)
            self.assertEqual(multi[2, 0], 1.1)
            self.assertEqual(multi[3, 0], "fish")
            self.assertEqual(multi.value, [(True,), (1,), (1.1,), ("fish",)])

            multi = XLOPER((True, 1, 1.1, "fish"))

            self.assertEqual(multi[0, 0], True)
            self.assertEqual(multi[0, 1], 1)
            self.assertEqual(multi[0, 2], 1.1)
            self.assertEqual(multi[0, 3], "fish")
            self.assertEqual(multi.value, [(True, 1, 1.1, "fish")])

            multi = XLOPER([[True, 1, 1.1, "fish"]])

            self.assertEqual(multi[0, 0], True)
            self.assertEqual(multi[0, 1], 1)
            self.assertEqual(multi[0, 2], 1.1)
            self.assertEqual(multi[0, 3], "fish")
            self.assertEqual(multi.value, [(True, 1, 1.1, "fish")])

            self.assertEqual(XLOPER([]).size, (0,0))
            self.assertEqual(XLOPER([()]).size, (1,0))

        def test_MultiString(self):
            """ensure we don't treat strings as sequences"""
            multi = XLOPER(["abc", "def"])
            self.assertEqual(multi[0, 0], "abc")
            self.assertEqual(multi[1, 0], "def")


        def test_exceptions(self):
            # TODO convert this to with assertRaises(...): format? 
            self.assertRaises(TypeError, XLOPER, complex(1, 2))
            self.assertRaises(TypeError, iter, XLOPER(1))
            self.assertRaises(TypeError, XLOPER(123).__getitem__, (0, 0))
            self.assertRaises(TypeError, int, XLOPER([1]))
            self.assertRaises(TypeError, float, XLOPER([1]))

            multi = XLOPER([[1, 2], [3, 4]])
            self.assertRaises(IndexError, multi.__getitem__, (0, 2))
            self.assertRaises(IndexError, multi.__getitem__, (2, 2))
            self.assertRaises(IndexError, multi.__getitem__, (-1, 0))
            self.assertRaises(IndexError, multi.__getitem__, (0, -1))

            with self.assertRaises(TypeError):
                XLOPER(RuntimeError()).value

        def test_eq(self):
            self.assertEqual(XLOPER(1), 1)
            self.assertEqual(XLOPER(1.0), 1.0)
            self.assertEqual(XLOPER(TypeError()), XLOPER(TypeError()))

    return _base


class TestXLOPER(_make(XLOPER4)):
    def test_unicode(self):
        self.assertEqual(type(XLOPER4(u'fish').value), str)
        self.assertEqual(XLOPER4(u'fish').value, 'fish')

class TestXLOPER12(_make(XLOPER12)):
    def test_unicode(self):
        self.assertEqual(type(XLOPER12(u'fish').value), unicode)
        self.assertEqual(XLOPER12('fish').value, u'fish')
        self.assertEqual(unicode(XLOPER12(u'V\u00e4rmev\u00e4rden AB')), u'V\u00e4rmev\u00e4rden AB')

    def test_str_length(self):
        super(TestXLOPER12, self).test_str_length(max=0xFFFF)        
