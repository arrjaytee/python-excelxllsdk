"""
enhance setuptools to  allow installation of xll entry points to allow python
to implement xll functions directly
"""

import os
import sys
import logging

from textwrap import dedent
from pkg_resources import resource_filename, ensure_directory, iter_entry_points
from setuptools.command.easy_install import current_umask, chmod

logging.basicConfig(level=logging.INFO)
LOG = logging.getLogger(__name__)

# TODO this could/should be done 100% inside the when xlAutoOpen is called
#      so DllMain loads in the entry point, and xlAutoOpen etc are mapped
#      onto it.

EP_GROUP = "excel_addins"
TEMPLATE = dedent(
    """
    # EASY-INSTALL-ENTRY-SCRIPT: %(spec)r,%(group)r,%(name)r
    __requires__ = %(spec)r
    from pkg_resources import load_entry_point
    DllMain = load_entry_point(%(spec)r, %(group)r, %(name)r)
    """
).lstrip()


def write_script(script_dir, script_name, contents, mode):
    """
    create a script with the specified text
    """
    LOG.info("Installing %s script to %s", script_name, script_dir)
    target = os.path.join(script_dir, script_name)
    mask = current_umask()
    ensure_directory(target)
    if os.path.exists(target):
        os.unlink(target)
    _file = open(target, "w" + mode)
    _file.write(contents)
    _file.close()
    chmod(target, 0o777 - mask)


def get_script_args(spec, name):
    """
    Yield write_script() argument tuples for a distribution's entrypoints
    allowing for excel_addin entry points
    """
    LOG.debug({"spec": spec, "group": EP_GROUP, "name": name})
    script_text = TEMPLATE % {"spec": spec, "group": EP_GROUP, "name": name}
    if sys.platform == "win32":
        yield (name + "-script.py", script_text, "t")
        yield (name + ".xll", get_addin_loader(), "b")
    else:
        raise RuntimeError("unsupported platform")


def write_entry_points():
    """
    search for packages in the lib directory which specify excel
    entry points, and install scripts for those entry points
    """
    script_dir = os.path.join(sys.exec_prefix, "scripts")
    for entry_point in iter_entry_points(EP_GROUP):
        pkg_name = entry_point.dist.project_name
        pkg_version = entry_point.dist.version
        for name, text, mode in get_script_args(
            "%s==%s" % (pkg_name, pkg_version), entry_point.name
        ):
            write_script(script_dir, name, text, mode)


def get_addin_loader():
    """
    return the add-in loader xll
    """
    # could use find_loader?
    return open(resource_filename("exceltools", "_addin32.xll"), "rb").read()
