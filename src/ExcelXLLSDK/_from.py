from __future__ import absolute_import

from math import isnan
from collections import Sequence
from types import NoneType
from multimethod import multimethod
from datetime import datetime

from .gen.xltype import *
from .gen.xlerr import *
from .xltypes import from_value, _XLOPER, _EXCEL_TIME_ZERO

import sys

@multimethod(_XLOPER, int)
def from_value(self, value):
    self.val.w = value
    self.xltype = xltypeInt

@multimethod(_XLOPER, long)
def from_value(self, value):
    # shoul use max of sys.val.w
    if value > sys.maxint or value < -(sys.maxint - 1):
        raise TypeError('python long is too large for excel int')
    self.val.w = value
    self.xltype = xltypeInt

@multimethod(_XLOPER, bool)
def from_value(self, value):
    self.val.xbool = 1 if value else 0
    self.xltype = xltypeBool

@multimethod(_XLOPER, float)
def from_value(self, value):
    if isnan(value):
        self.val.err = xlerrNA
        self.xltype = xltypeErr
    else:
        self.val.num = value
        self.xltype = xltypeNum

@multimethod(_XLOPER, str)
@multimethod(_XLOPER, unicode)
def from_value(self, value):
    self._set_Str(value)

@multimethod(_XLOPER, datetime)
def from_value(self, value):
    delta = value - _EXCEL_TIME_ZERO
    self.val.num = float(delta.days) + float(delta.seconds) / 86400
    self.xltype = xltypeNum

@multimethod(_XLOPER, Sequence)
def from_value(self, value):
    if len([x for x in value if not isinstance(x, Sequence) or isinstance(x, basestring)]) == 0:
        self._set_Multi(len(value), max((len(x) for x in value)) if len(value) else 0, iter(value))
    else:
        self._set_Multi(len(value), 1, ((x,) for x in value))

@multimethod(_XLOPER, tuple)
def from_value(self, value):
    if value == ():
        self.xltype = xltypeMissing
    else:
        self._set_Multi(1, len(value), [value])

@multimethod(_XLOPER, NoneType)
def from_value(self, value):
    self.xltype = xltypeNil
