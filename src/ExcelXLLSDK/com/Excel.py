"""ensure type library wrapper for excel is generated and import symbols"""
import sys
from comtypes.client import GetModule as _GetModule

_module = _GetModule(("{00020813-0000-0000-C000-000000000046}", 1, 7)) # 1.7 is Excel 2010?
globals().update((
    (key, value) for key, value in _module.__dict__.iteritems() if not key.startswith('_')
))
__all__ = _module.__all__
