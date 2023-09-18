from __future__ import unicode_literals, annotations
from typing import TYPE_CHECKING, Any
import uno
import unohelper
import sys
import os

from com.sun.star.task import XJob

if TYPE_CHECKING:
    # just for design time
    from lo_pip.config import Config
    from lo_pip.install.install_pip import InstallPip
    from lo_pip.install.install_pkg import InstallPkg
    from lo_pip.oxt_logger import OxtLogger

implementation_name = "___lo_identifier___.___lo_implementation_name___"
implementation_services = ("com.sun.star.task.Job",)


class ___lo_implementation_name___(unohelper.Base, XJob):
    def __init__(self, ctx):
        self._this_pth = os.path.dirname(__file__)
        self._path_added = False
        # logger.debug("OooPipRunner Init")
        self.ctx = ctx
        self._user_path = ""
        try:
            user_path = self._get_user_profile_path(True, self.ctx)
            # logger.debug(f"Init: user_path: {user_path}")
            self._user_path = user_path
        except Exception as err:
            # logger.error(err)
            pass

        self._add_local_path_to_sys_path()

        # for pth in sys.path:
        #     logger.debug(f"Sys.path at Init: {pth}")
        self._logger = self._get_local_logger()
        self._logger.debug("Got OxtLogger instance")
        self._logger.debug("OooPipRunner Init Done")

    def _get_user_profile_path(self, as_sys_path: bool = True, ctx: Any = None) -> str:
        """
        Returns the path to the user profile directory.

        Args:
            as_sys_path (bool): If True, returns the path as a system path entry otherwise ``file:///`` format.
                Defaults to True.
        """
        if ctx is None:
            ctx = uno.getComponentContext()
        result = ctx.ServiceManager.createInstance(
            "com.sun.star.util.PathSubstitution"
        ).substituteVariables(  # type: ignore
            "$(user)", True
        )
        if as_sys_path:
            return uno.fileUrlToSystemPath(result)
        return result

    def _get_local_logger(self) -> OxtLogger:
        from lo_pip.oxt_logger import OxtLogger

        # if self._user_path:
        #     log_file = os.path.join(self._user_path, "py_runner.log")
        #     return OxtLogger(log_file=log_file, log_name=__name__)
        return OxtLogger(log_name=__name__)

    def _add_local_path_to_sys_path(self) -> None:
        # add the path of this module to the sys.path
        if not self._this_pth in sys.path:
            sys.path.append(self._this_pth)
            # logger.debug(f"{self._this_pth}: appended to sys.path")
        self._path_added = True

    def _remove_local_path_from_sys_path(self) -> None:
        # remove the path of this module from the sys.path
        if not self._path_added:
            return
        if self._this_pth in sys.path:
            sys.path.remove(self._this_pth)
            # logger.debug(f"{self._this_pth}: removed from sys.path")
        self._path_added = False

    def _add_py_pkgs_to_sys_path(self) -> None:
        pth = os.path.join(os.path.dirname(__file__), "py_pkgs.zip")
        if not pth in sys.path:
            self._logger.debug(f"{pth}: appended to sys.path")
            sys.path.append(pth)

    def execute(self, *args):
        # make sure our pythonpath is in sys.path
        self._logger.debug("OooPipRunner executing")
        # logger = None
        try:
            self._add_local_path_to_sys_path()
            self._add_py_pkgs_to_sys_path()

            if not TYPE_CHECKING:
                # run time
                # from lo_pip.input_output import file_util
                from lo_pip.config import Config

                cfg = Config()

                self._logger.debug("Imported Config")
                from lo_pip.install.install_pip import InstallPip

                self._logger.debug("Imported InstallPip")
                from lo_pip.install.install_pkg import InstallPkg

                self._logger.debug("Imported InstallPkg")
            else:
                # design time
                cfg = Config()

            self._logger.debug("Created config instance")
            if cfg.py_pkg_dir:
                # add package zip file to the sys.path
                pth = os.path.join(os.path.dirname(__file__), f"{cfg.py_pkg_dir}.zip")

                if os.path.exists(pth) and os.path.isfile(pth) and os.path.getsize(pth) > 0 and not pth in sys.path:
                    self._logger.debug(f"{pth} appended to sys.path")
                    sys.path.append(pth)
            # sys.path.insert(0, sys.path.pop(sys.path.index(pth)))

            pip_installer = InstallPip()
            self._logger.debug("Created InstallPip instance")
            if pip_installer.is_pip_installed():
                self._logger.info("Pip is already installed")
            else:
                self._logger.info("Pip is not installed. Attempting to install")
                pip_installer.install_pip()
                if pip_installer.is_pip_installed():
                    self._logger.info("Pip has been installed")
                else:
                    self._logger.info("Pip was not successfully installed")

            # install any packages that are not installed
            pkg_installer = InstallPkg()
            self._logger.debug("Created InstallPkg instance")
            pkg_installer.install()

            self._logger.info("OooPipRunner execute Done!")
        except Exception as err:
            if self._logger:
                self._logger.error(err)
        finally:
            self._remove_local_path_from_sys_path()
        return


g_TypeTable = {}
# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

# add the FormatFactory class to the implementation container,
# which the loader uses to register/instantiate the component.
g_ImplementationHelper.addImplementation(___lo_implementation_name___, implementation_name, implementation_services)
