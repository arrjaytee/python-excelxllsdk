from __future__ import absolute_import
import sys
import logging
import _winreg
import argparse
import os

_log = logging.getLogger(__name__)

# should make this a setuptools extension, and use the list of entry points
# in the current virtualenv...

_reg_key = r"Software\Microsoft\Office\%(version)s\Excel\Options"
# _reg_versions = ['11.0', '12.0', '14.0']
_reg_versions = ['14.0']


def _find_xll(name):
    return os.path.join(sys.prefix, 'scripts', name + '.xll')


def unregister(name, versions=_reg_versions):
    """remove an from excels startup list"""
    # need to add a _d if we are debugging?
    basename = "%s.xll" % name.replace('.', os.sep)
    _log.info('Looking to unregister %s' % basename)

    for version in versions:
        key_name = _reg_key % {'version': version}

        try:
            key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER,
                                  key_name, 0,
                                  _winreg.KEY_READ | _winreg.KEY_WRITE)
        except WindowsError:
            _log.debug('Could not find %s' % key_name)
            continue

        # get count of keys to consder
        _, count, _ = _winreg.QueryInfoKey(key)

        # loop over all values in the register
        names = []
        for i in xrange(count):
            name, value, type = _winreg.EnumValue(key, i)

            if type != _winreg.REG_SZ or not name.startswith('OPEN'):
                continue

            # remove quotes
            value = value.strip('"')

            # clear up anything with _xltypes.pyd, what about _xltypes_d.pyd?
            if value.endswith(basename):
                names.append(name)

        for name in names:
            _log.info(
                "ExcelXLLSDK.register.unregister: %s\\%s = %s" % (key_name,
                                                                  name,
                                                                  value))
            _winreg.DeleteValue(key, name)


def register(name, versions=_reg_versions):
    filename = _find_xll(name)

    for version in versions:
        key_name = _reg_key % {'version': version}

        try:
            key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER,
                                  key_name, 0,
                                  _winreg.KEY_READ | _winreg.KEY_SET_VALUE)
        except WindowsError:
            _log.debug('Could not find %s' % key_name)
            continue

        # get count of keys to consder
        _, count, _ = _winreg.QueryInfoKey(key)

        # get a list of all open keys
        names = []
        for i in xrange(count):
            name, value, typ = _winreg.EnumValue(key, i)

            if name.startswith('OPEN'):
                names.append(name)

        # figure a new key for us
        i = 0
        name = 'OPEN'
        while name in names:
            i = i + 1
            name = "OPEN%d" % i

        _log.info("ExcelXLLSDK.register.register: %s\\%s = %s" % (key_name,
                                                                  name,
                                                                  filename))
        _winreg.SetValueEx(key, name, 0, _winreg.REG_SZ, '"%s"' % filename)


def UseCommandLine():
    parser = argparse.ArgumentParser(
        description='Register or Unregister a Python XLL Module',
        prefix_chars='-/'
    )
    parser.add_argument('--register',
                        action='store_true', help='Register XLL')
    parser.add_argument('--unregister',
                        action='store_true', help='Unregister XLL')
    parser.add_argument('modules', nargs=argparse.REMAINDER, metavar='MODULES')

    args = parser.parse_args(sys.argv[1:])

    if args.register:
        for module in args.modules:
            register(module)
    if args.unregister:
        for module in args.modules:
            unregister(module)

if __name__ == '__main__':
    UseCommandLine()
