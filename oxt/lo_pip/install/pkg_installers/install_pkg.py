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


# https://docs.python.org/3.8/library/importlib.metadata.html#module-importlib.metadata
# https://stackoverflow.com/questions/44210656/how-to-check-if-a-module-is-installed-in-python-and-if-not-install-it-within-t

# https://stackoverflow.com/search?q=%5Bpython%5D+run+subprocess+without+popup+terminal
# silent subprocess
if os.name == "nt":
    STARTUP_INFO = subprocess.STARTUPINFO()
    STARTUP_INFO.dwFlags |= subprocess.STARTF_USESHOWWINDOW
else:
    STARTUP_INFO = None


class InstallPkg:
    """Install pip packages."""

    def __init__(self) -> None:
        self._config = Config()
        self.path_python = Path(self._config.python_path)
        self.ver_rules = VerRules()
        self._logger = self._get_logger()

    def _get_logger(self) -> OxtLogger:
        return OxtLogger(log_name=__name__)

    def _get_package_version(self, package_name: str) -> str:
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

    def _cmd_pip(self, *args: str) -> List[str]:
        cmd: List[str] = [str(self.path_python), "-m", "pip", *args]
        return cmd

    def _install_pkg(self, pkg: str, ver: str) -> None:
        """
        Install a package.

        Args:
            pkg (str): The name of the package to install.
            ver (str): The version of the package to install.
        """
        auto_target = False
        if self.config.auto_install_in_site_packages:
            if self.config.site_packages:
                auto_target = True
            else:
                self._logger.debug(
                    "auto_install_in_site_packages is set to True; However, No site-packages directory set in configuration. site_packages value should be set in lo_pip.config.py"
                )
                self._logger.debug(
                    "Ignoring auto_install_in_site_packages and continuing to install in user directory via pip --user"
                )
        if auto_target:
            cmd = ["install", "--upgrade", f"--target={self.config.site_packages}"]
        else:
            cmd = ["install", "--upgrade", "--user"]
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

    def _get_env(self) -> Dict[str, str]:
        """
        Gets Environment used for subprocess.
        """
        my_env = os.environ.copy()
        py_path = ""
        p_sep = ";" if os.name == "nt" else ":"
        for d in sys.path:
            py_path = py_path + d + p_sep
        my_env["PYTHONPATH"] = py_path
        return my_env

    def install(self, req: Dict[str, str] | None = None) -> None:
        """
        Install all the packages in the configuration if they are not already installed and meet requirements.

        Args:
            req (Dict[str, str] | None, optional): The requirements to install.
                If omitted then requirements from config are used. Defaults to None.

        Returns:
            None:
        """
        self._logger.info("Installing packages…")
        req = req or self._config.requirements

        if not req:
            self._logger.warning("No packages to install.")
            return
        for name, ver in req.items():
            if name == "pip":
                continue
            if not ver:
                # set default version to >=0.0.0
                ver = "==*"
            pkg_ver = self._get_package_version(name)
            install = False
            rules = self.ver_rules.get_matched_rules(ver)
            if not pkg_ver:
                install = True
                self._logger.debug(f"Package {name} not installed. Setting Install flags.")

            else:
                self._logger.debug(f"Found Package {name} {pkg_ver} already installed ...")
                if not rules:
                    if pkg_ver:
                        self._logger.info(f"Package {name} {pkg_ver} already installed, no rules")
                    else:
                        self._logger.error(f"Unable to Install. Unable to find rules for {name} {ver}")
                    continue
                rules_pass = self.ver_rules.get_installed_is_valid_by_rules(rules=rules, check_version=pkg_ver)

                # for rule in rules:
                #     # rules are 'and' rules, so all rules must be true
                #     rules_pass = rules_pass and rule.get_version_is_valid(pkg_ver) == 0
                #     if rules_pass == False:
                #         break
                if rules_pass == False:
                    self._logger.info(
                        f"Package {name} {pkg_ver} already installed. It does not meet requirements specified by: {ver}, but will be upgraded."
                    )
                    install = True
                else:
                    self._logger.info(
                        f"Package {name} {pkg_ver} already installed; However, it does not need to be installed to meet constraints: {ver}. It will be skipped."
                    )
                    continue
            if install:
                ver_lst: List[str] = [rule.get_versions_str() for rule in rules]
                self._install_pkg(name, ",".join(ver_lst))
        self._logger.info("Installing packages Done!")

    @property
    def config(self) -> Config:
        return self._config
