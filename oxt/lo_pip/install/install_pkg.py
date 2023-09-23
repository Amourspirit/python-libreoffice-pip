from __future__ import annotations
import os

import subprocess
from typing import Dict


from pathlib import Path

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

    def install(self, req: Dict[str, str] | None = None) -> None:
        """
        Install all the packages in the configuration if they are not already installed and meet requirements.

        Args:
            req (Dict[str, str] | None, optional): The requirements to install.
                If omitted then requirements from config are used. Defaults to None.

        Returns:
            None:
        """
        if self._config.is_flatpak:
            self._logger.info("Flatpak detected, installing packages via Flatpak installer")
            self._install_flatpak(req=req)
            return

        self._logger.info("Installing packages via default installer")
        self._install_default(req=req)

    def _install_default(self, req: Dict[str, str] | None = None) -> None:
        from .pkg_installers.install_pkg import InstallPkg

        installer = InstallPkg()
        installer.install(req=req)

    def _install_flatpak(self, req: Dict[str, str] | None = None) -> None:
        from .pkg_installers.install_pkg_flatpak import InstallPkgFlatpak

        installer = InstallPkgFlatpak()
        installer.install(req=req)
