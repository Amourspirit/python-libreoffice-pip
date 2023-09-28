from __future__ import annotations
import os

import subprocess
from typing import Dict
from pathlib import Path
from importlib.metadata import PackageNotFoundError, version

from ..config import Config
from ..ver.rules.ver_rules import VerRules
from ..oxt_logger import OxtLogger


# https://docs.python.org/3.8/library/importlib.metadata.html#module-importlib.metadata
# https://stackoverflow.com/questions/44210656/how-to-check-if-a-module-is-installed-in-python-and-if-not-install-it-within-t

# https://stackoverflow.com/search?q=%5Bpython%5D+run+subprocess+without+popup+terminal
# silent subprocess
if os.name == "nt":
    _si = subprocess.STARTUPINFO()
    _si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
else:
    _si = None


class InstallPkg:
    """Install pip packages."""

    def __init__(self) -> None:
        self._config = Config()
        self.path_python = Path(self._config.python_path)
        self.ver_rules = VerRules()
        self._logger = OxtLogger(log_name=__name__)

    def install(self, req: Dict[str, str] | None = None, force: bool = False) -> None:
        """
        Install all the packages in the configuration if they are not already installed and meet requirements.

        Args:
            req (Dict[str, str] | None, optional): The requirements to install.
                If omitted then requirements from config are used. Defaults to None.
            force (bool, optional): Force install even if package is already installed. Defaults to False.

        Returns:
            None:
        """
        if self._config.is_flatpak:
            self._logger.info("Flatpak detected, installing packages via Flatpak installer")
            self._install_flatpak(req=req, force=force)
            return

        self._logger.info("Installing packages via default installer")
        self._install_default(req=req, force=force)

    def _install_default(self, req: Dict[str, str] | None, force: bool) -> None:
        from .pkg_installers.install_pkg import InstallPkg

        installer = InstallPkg()
        installer.install(req=req, force=force)

    def _install_flatpak(self, req: Dict[str, str] | None, force: bool) -> None:
        from .pkg_installers.install_pkg_flatpak import InstallPkgFlatpak

        installer = InstallPkgFlatpak()
        installer.install(req=req, force=force)

    def get_package_version(self, package_name: str) -> str:
        """
        Get the version of a package.

        Args:
            package_name (str): The name of the package such as ``verr``

        Returns:
            str: The version of the package or an empty string if the package is not installed.
        """
        try:
            return version(package_name)
        except PackageNotFoundError:
            return ""
