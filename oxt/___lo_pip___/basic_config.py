from __future__ import annotations
from pathlib import Path
from typing import Dict
import json

from .meta.singleton import Singleton


class ConfigMeta(type):
    _instance = None

    def __call__(cls, *args, **kwargs):
        if cls._instance is None:
            root = Path(__file__).parent
            config_file = Path(root, "config.json")
            with open(config_file, "r") as file:
                data = json.load(file)

            cls._instance = super().__call__(**data)
        return cls._instance


class BasicConfig(metaclass=ConfigMeta):
    def __init__(self, **kwargs) -> None:
        self._py_pkg_dir = str(kwargs["py_pkg_dir"])
        self._lo_identifier = str(kwargs["lo_identifier"])
        self._lo_implementation_name = str(kwargs["lo_implementation_name"])
        self._zipped_preinstall_pure = bool(kwargs["zipped_preinstall_pure"])
        self._auto_install_in_site_packages = bool(kwargs["auto_install_in_site_packages"])
        self._install_wheel = bool(kwargs["install_wheel"])
        self._has_locals = bool(kwargs["has_locals"])
        if "requirements" not in kwargs:
            kwargs["requirements"] = {}
        self._requirements: Dict[str, str] = dict(**kwargs["requirements"])

    # region Properties
    @property
    def auto_install_in_site_packages(self) -> bool:
        """
        Gets the flag indicating if packages are installed in the site-packages directory set in this config.

        The value for this property can be set in pyproject.toml (tool.oxt.config.auto_install_in_site_packages)

        If this is set to ``True`` then packages will be installed in the site-packages directory if this config has the value set.

        Flatpak ignores this setting and always installs packages in the site-packages directory determined in this config.

        Note:
            When running in a dev container (Codespace), this value is always set to ``True``.
        """
        return self._auto_install_in_site_packages

    @property
    def has_locals(self) -> bool:
        """
        Gets the flag indicating if the extension has local pip files to install.
        """
        return self._has_locals

    @property
    def install_wheel(self) -> bool:
        """
        Gets the flag indicating if wheel should be installed.
        """
        return self._install_wheel

    @property
    def lo_identifier(self) -> str:
        """
        Gets the LibreOffice identifier, such as, ``org.openoffice.extensions.ooopip``

        The value for this property can be set in pyproject.toml (tool.oxt.token.lo_identifier)
        """
        return self._lo_identifier

    @property
    def lo_implementation_name(self) -> str:
        """
        Gets the LibreOffice implementation name, such as ``OooPipRunner``

        The value for this property can be set in pyproject.toml (tool.oxt.token.lo_implementation_name)
        """
        return self._lo_implementation_name

    @property
    def py_pkg_dir(self) -> str:
        """
        The value for this property can be set in pyproject.toml (tool.oxt.config.py_pkg_dir)
        """
        return self._py_pkg_dir

    @property
    def requirements(self) -> Dict[str, str]:
        """
        Gets the set of requirements.

        The value for this property can be set in pyproject.toml (tool.oxt.token.requirements)

        The key is the name of the package and the value is the version number.
        Example: {"requests": ">=2.25.1"}
        """
        return self._requirements

    @property
    def zipped_preinstall_pure(self) -> bool:
        """
        Gets the flag indicating if pure python packages are be zipped.

        The value for this property can be set in pyproject.toml (tool.oxt.config.zip_preinstall_pure)

        If this is set to ``True`` then pure python packages will be zipped and installed as a zip file.
        """
        return self._zipped_preinstall_pure

    # endregion Properties
