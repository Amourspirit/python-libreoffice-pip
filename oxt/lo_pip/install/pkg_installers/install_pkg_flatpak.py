from __future__ import annotations
import os
import sys
import subprocess
from typing import Dict, List


# import pkg_resources
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

from ...config import Config
from ...ver.rules.ver_rules import VerRules
from ...oxt_logger import OxtLogger
from .install_pkg import InstallPkg
from .install_pkg import STARTUP_INFO


class InstallPkgFlatpak(InstallPkg):
    """Install pip packages for flatpak."""

    def _get_logger(self) -> OxtLogger:
        return OxtLogger(log_name=__name__)

    def _install_pkg(self, pkg: str, ver: str) -> None:
        """
        Install a package.

        Args:
            pkg (str): The name of the package to install.
            ver (str): The version of the package to install.
        """

        if not self.config.site_packages:
            self._logger.error(
                "No site-packages directory set in configuration. site_packages value should be set in lo_pip.config.py"
            )
            return
        cmd = ["install", "--upgrade", f"--target={self.config.site_packages}"]
        pkg_cmd = f"{pkg}{ver}" if ver else pkg
        cmd = self._cmd_pip(*[*cmd, pkg_cmd])
        self._logger.debug(f"Running command {cmd}")
        self._logger.info(f"Installing package {pkg}")
        msg = f"Pip Install - Upgrading success for: {pkg_cmd}"
        err_msg = f"Pip Install - Upgrading failed for: {pkg_cmd}"
        if STARTUP_INFO:
            process = subprocess.run(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=self._get_env(), startupinfo=STARTUP_INFO
            )
        else:
            process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=self._get_env())
        if process.returncode == 0:
            self._logger.info(msg)
        else:
            self._logger.error(err_msg)
        return
