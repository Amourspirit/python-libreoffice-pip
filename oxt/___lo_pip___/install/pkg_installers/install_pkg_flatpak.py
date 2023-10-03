from __future__ import annotations
import subprocess


# import pkg_resources
from ...oxt_logger import OxtLogger
from .install_pkg import InstallPkg
from .install_pkg import STARTUP_INFO


class InstallPkgFlatpak(InstallPkg):
    """Install pip packages for flatpak."""

    def _get_logger(self) -> OxtLogger:
        return OxtLogger(log_name=__name__)

    def _install_pkg(self, pkg: str, ver: str, force: bool) -> bool:
        """
        Install a package.

        Args:
            pkg (str): The name of the package to install.
            ver (str): The version of the package to install.
            force (bool): Force install even if package is already installed.

        Returns:
            bool: True if successful, False otherwise.
        """

        if not self.config.site_packages:
            self._logger.error(
                "No site-packages directory set in configuration. site_packages value should be set in lo_pip.config.py"
            )
            return False
        cmd = ["install"]
        if force:
            cmd.append("--force-reinstall")
        elif self.flag_upgrade:
            cmd.append("--upgrade")

        cmd.append(f"--target={self.config.site_packages}")

        pkg_cmd = f"{pkg}{ver}" if ver else pkg
        cmd = self._cmd_pip(*[*cmd, pkg_cmd])
        self._logger.debug(f"Running command {cmd}")
        self._logger.info(f"Installing package {pkg}")
        if self._flag_upgrade:
            msg = f"Pip Install - Upgrading success for: {pkg_cmd}"
            err_msg = f"Pip Install - Upgrading failed for: {pkg_cmd}"
        else:
            msg = f"Pip Install success for: {pkg_cmd}"
            err_msg = f"Pip Install failed for: {pkg_cmd}"
        if STARTUP_INFO:
            process = subprocess.run(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=self._get_env(), startupinfo=STARTUP_INFO
            )
        else:
            process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=self._get_env())
        if process.returncode == 0:
            self._logger.info(msg)
            return True
        else:
            self._logger.error(err_msg)
        return False
