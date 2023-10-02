# region imports
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
    from ___lo_pip___.config import Config
    from ___lo_pip___.install.install_pip import InstallPip
    from ___lo_pip___.install.install_pkg import InstallPkg
    from ___lo_pip___.oxt_logger import OxtLogger
    from ___lo_pip___.lo_util import Session, RegisterPathKind, UnRegisterPathKind
    from ___lo_pip___.lo_util.util import Util
    from ___lo_pip___.install.requirements_check import RequirementsCheck
else:
    RegisterPathKind = object
    UnRegisterPathKind = object

# endregion imports

# region Constants

implementation_name = "___lo_identifier___.___lo_implementation_name___"
implementation_services = ("com.sun.star.task.Job",)

# endregion Constants


# region XJob
class ___lo_implementation_name___(unohelper.Base, XJob):
    """Python UNO Component that implements the com.sun.star.task.Job interface."""

    # region Init

    def __init__(self, ctx):
        self._this_pth = os.path.dirname(__file__)
        self._path_added = False
        self._added_packaging = False
        # logger.debug("___lo_implementation_name___ Init")
        self.ctx = ctx
        self._user_path = ""
        with contextlib.suppress(Exception):
            user_path = self._get_user_profile_path(True, self.ctx)
            # logger.debug(f"Init: user_path: {user_path}")
            self._user_path = user_path
        self._add_local_path_to_sys_path()
        if not TYPE_CHECKING:
            # run time
            from ___lo_pip___.config import Config
            from ___lo_pip___.lo_util import Util

            # from ___lo_pip___.install.requirements_check import RequirementsCheck

            from ___lo_pip___.lo_util import (
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
        self._logger = self._get_local_logger()
        self._logger.debug("Got OxtLogger instance")

        # create an environment variable for the log file path.
        # This is used to give end users another way to find the log file via python.
        # Environment variable something like: ORG_OPENOFFICE_EXTENSIONS_OOOPIP_LOG_FILE, this will change for your extension.
        # The variable is determined by the lo_identifier in the pyproject.toml file, tool.oxt.token section.
        # To get the log file path in python: os.environ["ORG_OPENOFFICE_EXTENSIONS_OOOPIP_LOG_FILE"]
        # A Log file will only be created if log_file is set and log_level is not NONE, set in pyproject.toml file, tool.oxt.token section.
        if self._logger.log_file:
            log_env_name = self._config.lo_identifier.upper().replace(".", "_") + "_LOG_FILE"
            self._logger.debug(f"Log Path Environment Name: {log_env_name}")
            os.environ[log_env_name] = self._logger.log_file

        self._session = Session()
        self._logger.debug("___lo_implementation_name___ Init Done")

        self._add_py_req_pkgs_to_sys_path()
        if not TYPE_CHECKING:
            # run time
            # must be after self._add_py_req_pkgs_to_sys_path()
            from ___lo_pip___.install.requirements_check import RequirementsCheck
        self._requirements_check = RequirementsCheck()

    # endregion Init

    # region execute

    def execute(self, *args):
        # make sure our pythonpath is in sys.path
        self._logger.debug("___lo_implementation_name___ executing")
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

            if self._config.log_level < 20:  # Less than INFO
                self._show_extra_debug_info()
                # self._config.extension_info.log_extensions(self._logger)

            requirements_met = False
            if self._requirements_check.check_requirements() is True:
                # This will be True if all requirements in tool.oxt.requirements of pyproject.toml are met.
                # Also, This speeds up the loading of the extension considerably if no requirements need installing.

                if not self._config.has_locals:
                    requirements_met = True

            if requirements_met:
                self._logger.debug("Requirements are met. Nothing more to do.")
                return

            if not TYPE_CHECKING:
                # run time
                from ___lo_pip___.install.install_pip import InstallPip

                self._logger.debug("Imported InstallPip")
                from ___lo_pip___.install.install_pkg import InstallPkg

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
                if not pip_installer.is_internet:
                    self._logger.error("No internet connection!")
                    return
                pip_installer.install_pip()
                if pip_installer.is_pip_installed():
                    self._logger.info("Pip has been installed")
                else:
                    self._logger.info("Pip was not successfully installed")
                    return

            # install wheel if needed
            self._install_wheel()
            # install any packages that are not installed
            if self._config.has_locals:
                self._install_locals()
            pkg_installer = InstallPkg()
            self._logger.debug("Created InstallPkg instance")
            pkg_installer.install()

            self._logger.info(f"{self._config.lo_implementation_name} execute Done!")
        except Exception as err:
            if self._logger:
                self._logger.error(err)
        finally:
            self._remove_local_path_from_sys_path()
            self._remove_py_req_pkgs_from_sys_path()
            self._log_ex_time(start_time)

    # endregion execute

    # region Destructor
    def __del__(self):
        if self._added_packaging:
            if "packaging" in sys.modules:
                # clean up by removing the packaging module from sys.modules
                del sys.modules["packaging"]
        if "___lo_pip___" in sys.modules:
            # clean up by removing the ___lo_pip___ module from sys.modules
            del sys.modules["___lo_pip___"]

    # endregion Destructor

    # region Install

    def _install_wheel(self) -> None:
        if not self._config.install_wheel:
            self._logger.debug("Install wheel is set to False. Skipping wheel installation.")
            return
        self._logger.debug("Install wheel is set to True. Installing wheel.")
        try:
            from ___lo_pip___.install.install_wheel import InstallWheel

            installer = InstallWheel()
            installer.install()
        except Exception as err:
            self._logger.error(f"Unable to install wheel: {err}", exc_info=True)
            return
        self._logger.debug("Install wheel done.")

    # endregion Install

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
        pth = Path(os.path.dirname(__file__), f"{self._config.py_pkg_dir}.zip")
        if not pth.exists():
            return
        result = self._session.register_path(pth, True)
        self._log_sys_path_register_result(pth, result)

    def _add_pure_pkgs_to_sys_path(self) -> None:
        pth = Path(os.path.dirname(__file__), "pure.zip")
        if not pth.exists():
            self._logger.debug(f"pure.zip not found.")
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
            if result == RegisterPathKind.REGISTERED:
                self._added_packaging = True

    def _remove_py_req_pkgs_from_sys_path(self) -> None:
        pth = Path(os.path.dirname(__file__), f"req_{self._config.py_pkg_dir}.zip")
        result = self._session.unregister_path(pth)
        self._log_sys_path_unregister_result(pth, result)

    def _add_site_package_dir_to_sys_path(self) -> None:
        if self._config.is_shared_installed or self._config.is_bundled_installed:
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

    # region other methods

    def _log_ex_time(self, start_time: float, msg: str = "") -> None:
        if not self._logger:
            return
        end_time = time.time()
        total_time = end_time - start_time
        self._logger.info(f"{self._config.lo_implementation_name} execution time: {total_time:.3f} seconds")

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

    # endregion other methods

    # region install local
    def _install_locals(self) -> None:
        """Pip installs any ``.whl`` or ``.tar.gz`` files in the ``locals`` directory."""
        if not self._config.has_locals:
            self._logger.debug("Install local is set to False. Skipping local installation.")
            return
        self._logger.debug("Install local is set to True. Installing local packages.")
        try:
            from ___lo_pip___.install.install_pkg_local import InstallPkgLocal

            installer = InstallPkgLocal()
            installer.install()
        except Exception as err:
            self._logger.error(f"Unable to install local packages: {err}", exc_info=True)
            return
        self._logger.debug("Install local done.")

    # endregion install local
    # region Logging

    def _get_local_logger(self) -> OxtLogger:
        from ___lo_pip___.oxt_logger import OxtLogger

        # if self._user_path:
        #     log_file = os.path.join(self._user_path, "py_runner.log")
        #     return OxtLogger(log_file=log_file, log_name=__name__)
        return OxtLogger(log_name=__name__)

    # endregion Logging

    # region Debug

    def _show_extra_debug_info(self):
        self._logger.debug(f"Config Package Location: {self._config.package_location}")
        self._logger.debug(f"Config Is User Installed: {self._config.is_user_installed}")
        self._logger.debug(f"Config Is Share Installed: {self._config.is_shared_installed}")
        self._logger.debug(f"Config Is Bundle Installed: {self._config.is_bundled_installed}")

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

    # endregion Debug


# endregion XJob

# region Implementation

g_TypeTable = {}
# python loader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

# add the FormatFactory class to the implementation container,
# which the loader uses to register/instantiate the component.
g_ImplementationHelper.addImplementation(___lo_implementation_name___, implementation_name, implementation_services)

# endregion Implementation
