from __future__ import annotations
from typing import List, Dict
import sys

from .py_package import PyPackage
from .package_config import PackageConfig
from ...config import Config
from ...oxt_logger import OxtLogger
from ...ver.rules.ver_rules import VerRules


class Packages:
    """Manages rules for Python Packages"""

    def __init__(self) -> None:
        """
        Initialize PackageRules
        """
        self._config = Config()
        self._log = OxtLogger(log_name=self.__class__.__name__)
        self._packages: List[PyPackage] = []
        self._py_ver = self._get_py_ver()
        self._load_packages()

    def _get_py_ver(self) -> str:
        # This is here for easier testing. It can be mocked in tests.
        return f"{sys.version_info[0]}.{sys.version_info[1]}.{sys.version_info[2]}"

    def _load_packages(self) -> None:
        """
        Load rules from config
        """
        ver_rules = VerRules()

        def is_valid(rule: PyPackage) -> bool:
            nonlocal ver_rules

            if rule.python_versions:
                is_package_valid_python = any(
                    ver_rules.get_installed_is_valid(f">={ver}", self._py_ver) for ver in rule.python_versions
                )
                if is_package_valid_python:
                    self._log.debug(
                        "Python version %s for %s is valid for python requirement", self._py_ver, rule.name
                    )
                else:
                    self._log.debug(
                        "Python version %s for %s is NOT valid for python requirement %s", self._py_ver, rule.name
                    )
                    return False

            if self._config.is_win:
                platform = "win"
            elif self._config.is_mac:
                platform = "mac"
            elif self._config.is_flatpak:
                platform = "flatpak"
            elif self._config.is_snap:
                platform = "snap"
            else:
                platform = "linux"
            if rule.is_ignored_platform(platform):
                return False
            return rule.is_platform(platform)

        pkg_cfg = PackageConfig()
        for rule in pkg_cfg.py_packages:
            gi = PyPackage.from_dict(**rule)
            if is_valid(gi):
                self._log.debug("Adding rule: %s}", gi)
                self.add_pkg(gi)
            else:
                self._log.debug("Ignoring rule: %s}", gi)

    def __len__(self) -> int:
        return len(self._packages)

    def __contains__(self, pkg: PyPackage) -> bool:
        return pkg in self._packages

    # region Methods

    def get_index(self, pkg: PyPackage) -> int:
        """
        Get index of Package

        Args:
            pkg (PyPackage): Rule to get index

        Returns:
            int: Index of rule
        """
        return self._packages.index(pkg)

    def add_pkg(self, pkg: PyPackage) -> None:
        """
        Add Package

        Args:
            pkg (PyPackage): Rule to register
        """
        if pkg in self._packages:
            self._log.debug("add_pkg() Rule Already added: %s", pkg)
            return
        self._log.debug("add_pkg() Adding Rule %s", self.__class__.__name__, pkg)
        self._add_pkg(pkg=pkg)

    def remove_package(self, pkg: PyPackage):
        """
        Removes Package

        Args:
            pkg (PyPackage): pkg to remove

        Raises:
            ValueError: If an error occurs
        """
        try:
            self._packages.remove(pkg)
            self._log.debug("remove_package() Removed rule: %s,", pkg)
        except ValueError as e:
            msg = f"{self.__class__.__name__}.unregister_rule() Unable to unregister rule."
            self._log.error(msg)
            raise ValueError(msg) from e

    def _add_pkg(self, pkg: PyPackage):
        self._packages.append(pkg)

    def __repr__(self) -> str:
        return "<Packages()>"

    def to_dict(self) -> Dict[str, str]:
        """
        Convert to dict with Name as key, restriction and version as value.


        Returns:
            dict: Dict representation of the object such as ``{"verr": ">=1.0.0", "requests": "==2.0.0"}``
        """
        result = {}
        for pkg in self.packages:
            result[pkg.name] = f"{pkg.restriction}{pkg.version}"
        return result

    # endregion Methods

    # region Properties
    @property
    def packages(self) -> List[PyPackage]:
        """
        Get all rules

        Returns:
            List[PyPackage]: List of all rules
        """
        return self._packages

    # endregion Properties
