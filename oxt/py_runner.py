from __future__ import unicode_literals, annotations
import contextlib
from typing import TYPE_CHECKING, Any
from pathlib import Path
import uno
import unohelper
import sys
import os
import time

from com.sun.star.task import XJob

if TYPE_CHECKING:
    # just for design time
    from lo_pip.config import Config
    from lo_pip.install.install_pip import InstallPip
    from lo_pip.install.install_pkg import InstallPkg
    from lo_pip.oxt_logger import OxtLogger
    from lo_pip.lo_util import Session, RegisterPathKind, UnRegisterPathKind
    from lo_pip.lo_util.util import Util
    from lo_pip.info import ExtensionInfo
else:
    RegisterPathKind = object
    UnRegisterPathKind = object

implementation_name = "___lo_identifier___.___lo_implementation_name___"
implementation_services = ("com.sun.star.task.Job",)


class ___lo_implementation_name___(unohelper.Base, XJob):
    def __init__(self, ctx):
        self._this_pth = os.path.dirname(__file__)
        self._path_added = False
        # logger.debug("OooPipRunner Init")
        self.ctx = ctx
        self._user_path = ""
        with contextlib.suppress(Exception):
            user_path = self._get_user_profile_path(True, self.ctx)
            # logger.debug(f"Init: user_path: {user_path}")
            self._user_path = user_path
        self._add_local_path_to_sys_path()
        if not TYPE_CHECKING:
            # run time
            from lo_pip.config import Config
            from lo_pip.lo_util import Util
            from lo_pip.info import ExtensionInfo

            from lo_pip.lo_util import (
                Session,
                RegisterPathKind as InitRegisterPathKind,
                UnRegisterPathKind as InitUnRegisterPathKind,
            )

            global RegisterPathKind, UnRegisterPathKind
            RegisterPathKind = InitRegisterPathKind
            UnRegisterPathKind = InitUnRegisterPathKind

        # design time
        self._config = Config()
        self._util = Util()
        self._extension_info = ExtensionInfo()
        self._logger = self._get_local_logger()
        self._logger.debug("Got OxtLogger instance")
        self._session = Session()
        self._logger.debug("OooPipRunner Init Done")

        self._add_py_req_pkgs_to_sys_path()

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
        return uno.fileUrlToSystemPath(result) if as_sys_path else result

    def _get_local_logger(self) -> OxtLogger:
        from lo_pip.oxt_logger import OxtLogger

        # if self._user_path:
        #     log_file = os.path.join(self._user_path, "py_runner.log")
        #     return OxtLogger(log_file=log_file, log_name=__name__)
        return OxtLogger(log_name=__name__)

    # region Register/Unregister sys paths

    def _add_local_path_to_sys_path(self) -> None:
        # add the path of this module to the sys.path
        if self._this_pth not in sys.path:
            sys.path.append(self._this_pth)
            # logger.debug(f"{self._this_pth}: appended to sys.path")
        self._path_added = True

    def _remove_local_path_from_sys_path(self) -> None:
        # remove the path of this module from the sys.path
        if not self._path_added:
            return
        if self._this_pth in sys.path:
            sys.path.remove(self._this_pth)
            self._logger.debug(f"sys.path removed: {self._this_pth}")
        self._path_added = False

    def _add_py_pkgs_to_sys_path(self) -> None:
        # pth = os.path.join(os.path.dirname(__file__), "py_pkgs.zip")
        pth = Path(os.path.dirname(__file__), f"{self._config.py_pkg_dir}.zip")
        if not pth.exists():
            return
        result = self._session.register_path(pth, True)
        self._log_sys_path_register_result(pth, result)

    def _add_pure_pkgs_to_sys_path(self) -> None:
        # pth = os.path.join(os.path.dirname(__file__), "py_pkgs.zip")
        pth = Path(os.path.dirname(__file__), "pure.zip")
        if not pth.exists():
            self._logger.debug(f"pure.zip not found: {pth}")
            return
        result = self._session.register_path(pth, True)
        self._log_sys_path_register_result(pth, result)

    def _add_py_req_pkgs_to_sys_path(self) -> None:
        pth = Path(os.path.dirname(__file__), f"req_{self._config.py_pkg_dir}.zip")
        if not pth.exists():
            return
        # should be only LibreOffice on Windows needs packaging
        try:
            self._logger.debug("Importing packaging")
            import packaging  # noqa: F401

            self._logger.debug("packaging imported")
        except ModuleNotFoundError:
            self._logger.debug("packaging not found. Adding to sys.path")
            result = self._session.register_path(pth, True)
            self._log_sys_path_register_result(pth, result)

    def _remove_py_req_pkgs_from_sys_path(self) -> None:
        pth = Path(os.path.dirname(__file__), f"req_{self._config.py_pkg_dir}.zip")
        result = self._session.unregister_path(pth)
        self._log_sys_path_unregister_result(pth, result)

    def _add_site_package_dir_to_sys_path(self) -> None:
        if self._config.is_all_users:
            self._logger.debug("All users, not adding site-packages to sys.path")
            return
        if not self._config.site_packages:
            return
        result = self._session.register_path(self._config.site_packages, True)
        self._log_sys_path_register_result(self._config.site_packages, result)

    def _log_sys_path_register_result(self, pth: Path | str, result: RegisterPathKind) -> None:
        if not isinstance(pth, str):
            pth = str(pth)
        if result == RegisterPathKind.NOT_REGISTERED:
            if not pth:
                self._logger.debug("Path not registered. Can't register empty string")
            else:
                self._logger.debug(f"Path Not Registered, unknown reason: {pth}")
        elif result == RegisterPathKind.ALREADY_REGISTERED:
            self._logger.debug(f"Path already registered: {pth}")
        else:
            self._logger.debug(f"Path registered: {pth}")

    def _log_sys_path_unregister_result(self, pth: Path | str, result: UnRegisterPathKind) -> None:
        if not isinstance(pth, str):
            pth = str(pth)
        if result == UnRegisterPathKind.NOT_UN_REGISTERED:
            if not pth:
                self._logger.debug("Path not unregistered. Can't unregister empty string")
            else:
                self._logger.debug(f"Path Not unregistered, unknown reason: {pth}")
        elif result == UnRegisterPathKind.ALREADY_UN_REGISTERED:
            self._logger.debug(f"Path already unregistered: {pth}")
        else:
            self._logger.debug(f"Path unregistered: {pth}")

    # endregion Register/Unregister sys paths

    def execute(self, *args):
        # make sure our pythonpath is in sys.path
        self._logger.debug("OooPipRunner executing")
        start_time = time.time()
        # if args:
        #     self._logger.debug(f"args: {args}")

        # logger = None
        try:
            self._add_local_path_to_sys_path()
            self._add_py_pkgs_to_sys_path()
            self._add_py_req_pkgs_to_sys_path()
            self._add_pure_pkgs_to_sys_path()
            self._add_site_package_dir_to_sys_path()

            self._logger.debug(f"Config Package Location: {self._config.package_location}")
            self._logger.debug(f"Config Is All Users: {self._config.is_all_users}")

            self._logger.debug(f"Session - LibreOffice Share: {self._session.share}")
            self._logger.debug(f"Session - LibreOffice Share Python: {self._session.shared_py_scripts}")
            self._logger.debug(f"Session - LibreOffice Share Scripts: {self._session.shared_scripts}")
            self._logger.debug(f"Session - LibreOffice Username: {self._session.user_name}")
            self._logger.debug(f"Session - LibreOffice User Profile: {self._session.user_profile}")
            self._logger.debug(f"Session - LibreOffice User Scripts: {self._session.user_scripts}")

            self._logger.debug(f"Util.config - Module: {self._util.config('Module')}")
            self._logger.debug(f"Util.config - UserConfig: {self._util.config('UserConfig')}")
            self._logger.debug(f"Util.config - Config: {self._util.config('Config')}")
            self._logger.debug(f"Util.config - BasePathUserLayer: {self._util.config('BasePathUserLayer')}")
            self._logger.debug(f"Util.config - BasePathShareLayer: {self._util.config('BasePathShareLayer')}")

            self._extension_info.log_extensions()
            ext_info = self._extension_info.get_extension_info(id=self._config.lo_identifier)
            self._logger.debug(f"Extension Info: {ext_info}")

            if not TYPE_CHECKING:
                # run time
                from lo_pip.install.install_pip import InstallPip

                self._logger.debug("Imported InstallPip")
                from lo_pip.install.install_pkg import InstallPkg

            self._logger.debug("Created config instance")
            if self._config.py_pkg_dir:
                # add package zip file to the sys.path
                pth = os.path.join(os.path.dirname(__file__), f"{self._config.py_pkg_dir}.zip")

                if os.path.exists(pth) and os.path.isfile(pth) and os.path.getsize(pth) > 0 and pth not in sys.path:
                    self._logger.debug(f"sys.path appended: {pth}")
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
                    return

            # install any packages that are not installed
            pkg_installer = InstallPkg()
            self._logger.debug("Created InstallPkg instance")
            pkg_installer.install()

            self._logger.info("___lo_implementation_name___ execute Done!")
        except Exception as err:
            if self._logger:
                self._logger.error(err)
        finally:
            self._remove_local_path_from_sys_path()
            self._remove_py_req_pkgs_from_sys_path()
        end_time = time.time()
        if self._logger:
            self._logger.info(f"___lo_implementation_name___ execution time: {end_time - start_time} seconds")
        return


g_TypeTable = {}
# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

# add the FormatFactory class to the implementation container,
# which the loader uses to register/instantiate the component.
g_ImplementationHelper.addImplementation(___lo_implementation_name___, implementation_name, implementation_services)
