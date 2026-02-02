from __future__ import absolute_import

import os, os.path, sys, argparse
import pkg_resources
import logging
from exceltools.registry import _find_xll
from .process import ExcelPopen, excel_versions

def main(argv=None):
    """start and excel process and run a given entry point in that module"""

    # parse options?
    parser = argparse.ArgumentParser("invoke a python script or module within an excel process")
    parser.add_argument("--excel-version", dest='version',
                        action="store",
                        metavar="VERSION",
                        default='2010',
                        help="excel version to run",
                        choices=excel_versions.values() + excel_versions.keys()
                        )
    parser.add_argument("--no-entrypoints", dest='no_entrypoints',
                        action="store_true",
                        default=False,
                        help="do not load python entry points",
                        )
    parser.add_argument("--no-user", dest='no_user',
                        action='store_true',
                        default=False, help="do not import user in the excel process"
                        )
    parser.add_argument("-i", dest='interactive',
                        action='store_true',
                        default=False,
                        help="keep excel window visible, allow interaction when done"
                        )

    parser.add_argument("-c", dest='command',
                        nargs=argparse.REMAINDER,
                        metavar='CMD',
                        default=[],
                        help="invoke module as main script, all subsequent arguments are passed through"
                        )
    parser.add_argument("-m", dest='module',
                        nargs=argparse.REMAINDER,
                        default=[],
                        help="invoke module as main script, all subsequent arguments are passed through"
                        )
    parser.add_argument("script",
                        metavar="SCRIPT",
                        nargs=argparse.REMAINDER,
                        help="python script to run, all subsequent arguments are passed through"
                        )
    args = parser.parse_args(argv or sys.argv[1:])

    # TODO add a launch mode where we detach and get a new console? may need to signal to the addin C code
    ps = ExcelPopen(
        version=args.version,
        interactive=args.interactive,
        visible=args.interactive,
        shell=False,
        stdout=sys.stdout,
        stderr=sys.stderr,
        stdin=sys.stdin,
    )

    # somehow this logic could be in the argparser?
    if args.command:
        argv = ['-c'] + args.command
    elif args.module:
        argv = ['-m'] + args.module
    else:
        argv = args.script

        # if we specified an entry point, use that script instead
        if argv:
            entry_script = os.path.join(sys.prefix, "Scripts", args.script[0]+"-script.py")
            if os.path.exists(entry_script):
                argv[0] = entry_script

    if args.no_entrypoints:
        entry_points = [ '_pyxcel' ]
    else:
        entry_points = [
            ep.name for ep in pkg_resources.iter_entry_points('excel_addins') if not ep.name.startswith('_')
        ] + ['_pyxcel']

    for filename in map(_find_xll, entry_points):
        print 'RegisterXLL: ', filename
        if not ps.Application.RegisterXLL(filename):
            raise RuntimeError('could not load %s' % filename)

    macros = [ '_pyxcel_chdir("%s")' % os.path.normcase(os.getcwd()) ]
    if not args.no_user:
        macros.append('_pyxcel_import_user()')
    if argv:
        macros.append('_pyxcel_cmdline("%s")' % repr(argv).replace('"', '""'))

    try:
        for macro in macros:
            exitcode = ps.Application.ExecuteExcel4Macro(macro)
            if exitcode:
                raise SystemExit(exitcode)
    except:
        sys.excepthook(*sys.exc_info())
        raise
    finally:
        if args.interactive or not argv:
            ps.Application.Interactive = True
            ps.Application.DisplayAlerts = True
            ps.Application.Visible = True
            del ps.Application
        else:
            ps.quit()

        print 'waiting for excel process {:d} to finish'.format(ps.pid)
        ps.wait()

if __name__ == "__main__":
    main()