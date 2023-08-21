from __future__ import unicode_literals, annotations
from typing import TYPE_CHECKING
import uno
import unohelper
import sys
import os

from com.sun.star.task import XJob

if TYPE_CHECKING:
    # just for design time
    from pythonpath.lo_pip.config import Config
    from pythonpath.lo_pip.install.install_pip import InstallPip

implementation_name = "___lo_identifier___.___lo_implementation_name___"
implementation_services = ("com.sun.star.task.Job",)


class ___lo_implementation_name___(unohelper.Base, XJob):
    def __init__(self, ctx):
        self.ctx = ctx

    def execute(self, *args):
        # make sure our pythonpath is in sys.path
        pth = os.path.join(os.path.dirname(__file__), "py_pkgs.zip")
        if not pth in sys.path:
            sys.path.append(pth)
        return

        pth = os.path.join(os.path.dirname(__file__), "pythonpath")
        if os.path.exists(pth) and os.path.isdir(pth) and not pth in sys.path:
            sys.path.append(pth)

        if not TYPE_CHECKING:
            # run time
            from lo_pipe.config import Config
            from lo_pipe.install.install_pip import InstallPip

        cfg = Config()
        if cfg.py_pkg_dir:
            # add package zip file to the sys.path
            pth = os.path.join(os.path.dirname(__file__), f"{cfg.py_pkg_dir}.zip")

            if os.path.exists(pth) and os.path.isfile(pth) and os.path.getsize(pth) > 0 and not pth in sys.path:
                sys.path.append(pth)
        # sys.path.insert(0, sys.path.pop(sys.path.index(pth)))

        pip_installer = InstallPip()
        if not pip_installer.is_pip_installed():
            pip_installer.install_pip()

        return


g_TypeTable = {}
# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

# add the FormatFactory class to the implementation container,
# which the loader uses to register/instantiate the component.
g_ImplementationHelper.addImplementation(___lo_implementation_name___, implementation_name, implementation_services)
