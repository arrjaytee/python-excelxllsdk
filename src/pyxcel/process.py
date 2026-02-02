"""run an excel as a command line process"""
import logging

import os, os.path, subprocess,  time, logging
from ExcelXLLSDK.com import find_Application
from _winreg import OpenKey, QueryValueEx, HKEY_LOCAL_MACHINE

excel_versions = {
    '2003': '11.0',
    '2007': '12.0',
    '2010': '14.0'
}

_default_version = '14.0'

def _get_executable(version=_default_version):
    try:
        hkey = OpenKey(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Office\\%s\\Excel\\InstallRoot" % version)
    except WindowsError:
        raise RuntimeError("Could not find registry entries for Excel %s" % version)
    path, _ = QueryValueEx(hkey, "Path")
    path = os.path.join(path, 'EXCEL.EXE')
    os.stat(path)
    return path


def _get_installed_versions():
    for version in excel_versions.values():
        try:
            _ = _get_executable(version=version)
            yield version
        except RuntimeError:
            pass
        except WindowsError:
            pass

# TODO launch withouth inheriting console, we we make our own?
class ExcelPopen(subprocess.Popen):
    """create and maintain an excel process, provide access to the COM model"""
    def __init__(self, files=[], version=_default_version, safemode=False, visible=False, interactive=False, **kw):
        """initialize a private process for the given excel version"""

        version = excel_versions.get(version, version)

        executable = _get_executable(version=version)
        args = [executable]
        args.extend(files)

        startupinfo = subprocess.STARTUPINFO()

        if not visible:
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

        super(ExcelPopen, self).__init__(args, startupinfo=startupinfo, cwd=os.getcwd(), **kw)

        # allow time for the process to start
        for _ in xrange(0, 10):
            try:
                self.Application = find_Application(self.pid)
                break
            except RuntimeError:
                time.sleep(1)
                continue
        else:
            raise RuntimeError('could not find Excel Application object from windows')

        # setup application state to reflect this process
        self.Application.Visible = visible
        self.Application.DisplayAlerts = visible

    def quit(self):
        """shutdown the excel process cleanly using Application.Quit()"""
        self.Application.DisplayAlerts = False
        self.Application.Quit()
        del self.Application
        self.wait()

    def __del__(self):
        """terminate the process if it hasn't already been shutdown safely"""
        if self.returncode == None:
            self.terminate()