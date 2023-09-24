# coding: utf-8
from __future__ import annotations
from pathlib import Path
from typing import Dict
import json
import os
import sys
import platform

from .input_output import file_util


class ConfigMeta(type):
    _instance = None

    def __call__(cls, *args, **kwargs):
        if cls._instance is None:
            root = Path(__file__).parent
            config_file = Path(root, "config.json")
            if config_file.exists():
                with open(config_file, "r") as file:
                    data = json.load(file)
                    # logger.debug("Configuration: Loaded config.json")
            else:
                data = {
                    "url_pip": "https://bootstrap.pypa.io/get-pip.py",
                    "log_file": "lo_pip.log",
                    "log_name": "OOO PIP Installer",
                    "log_level": "INFO",
                    "log_format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
                    "py_pkg_dir": "py_pkgs",
                    "pip_wheel_url": "",
                    "lo_identifier": "",
                    "lo_implementation_name": "",
                    "requirements": {},
                    "zipped_preinstall_pure": False,
                    "auto_install_in_site_packages": False,
                }
                # logger.debug("Configuration: no config.json, using defaults")

            cls._instance = super().__call__(**data)
        return cls._instance


OS = platform.system()
IS_WIN = OS == "Windows"
IS_MAC = OS == "Darwin"


class Config(metaclass=ConfigMeta):
    """
    Singleton Configuration Class

    Generally speaking this class is only used internally.
    """

    def __init__(self, **kwargs):
        from .lo_util import Util

        log_file = str(kwargs["log_file"])
        if log_file:
            log_pth = Path(log_file)
            if not log_pth.is_absolute():
                log_file = Path(file_util.get_user_profile_path(True), log_pth)
        self._log_file = str(log_file)
        # self._log_file = "D:\\tmp\\log\\py_runner.log"
        # logger.debug("Config.__init__ called")
        self._url_pip = str(kwargs["url_pip"])
        self._pip_wheel_url = str(kwargs["pip_wheel_url"])
        self._log_name = str(kwargs["log_name"])
        self._log_format = str(kwargs["log_format"])
        self._py_pkg_dir = str(kwargs["py_pkg_dir"])
        self._lo_identifier = str(kwargs["lo_identifier"])
        self._lo_implementation_name = str(kwargs["lo_implementation_name"])
        self._zipped_preinstall_pure = bool(kwargs["zipped_preinstall_pure"])
        # auto_install_in_site_packages
        self._auto_install_in_site_packages = bool(kwargs["auto_install_in_site_packages"])
        if not self._auto_install_in_site_packages and os.getenv("DEV_CONTAINER", "") == "1":
            self._auto_install_in_site_packages = True
        if "requirements" not in kwargs:
            kwargs["requirements"] = {}
        self._requirements: Dict[str, str] = dict(**kwargs["requirements"])
        self._log_level = self._get_log_level(kwargs["log_level"])
        self._os = platform.system()
        self._is_win = self._os == "Windows"
        self._is_mac = self._os == "Darwin"
        self._is_app_image = False
        app_img = os.getenv("APPIMAGE", "")
        if app_img:
            self._is_app_image = True
        flat_pak_id = os.getenv("FLATPAK_ID", "")
        self._is_flatpak = flat_pak_id.lower() == "org.libreoffice.libreoffice"
        self._site_packages = ""
        util = Util()
        self._python_major_minor = self._get_python_major_minor()

        if self._is_win:
            self._python_path = Path(self.join(util.config("Module"), "python.exe"))
        elif self._is_mac:
            self._python_path = Path(self.join(util.config("Module"), "..", "Resources", "python"))
            self._site_packages = self._get_mac_site_packages_dir()
        elif self._is_app_image:
            self._python_path = Path(self.join(util.config("Module"), "python"))
            self._site_packages = self._get_default_site_packages_dir()
        else:
            self._python_path = Path(sys.executable)
            if self._is_flatpak:
                self._site_packages = self._get_flatpak_site_packages_dir()
            else:
                self._site_packages = self._get_default_site_packages_dir()

        # logger.debug("Config.__init__ completed")

    def join(self, *paths: str):
        return str(Path(paths[0]).joinpath(*paths[1:]))

    def _get_log_level(self, log_level: str | int) -> int:
        if isinstance(log_level, str):
            log_level = log_level.upper()
            if log_level == "NONE":
                return 0
            if log_level == "DEBUG":
                return 10
            elif log_level == "INFO":
                return 20
            elif log_level == "WARNING":
                return 30
            elif log_level == "ERROR":
                return 40
            elif log_level == "CRITICAL":
                return 50
            else:
                raise ValueError(f"Invalid log level: {log_level}")
        elif isinstance(log_level, int):
            return log_level
        else:
            raise TypeError(f"Invalid log level type: {type(log_level)}")

    def _get_python_major_minor(self) -> str:
        return f"{sys.version_info.major}.{sys.version_info.minor}"

    def _get_default_site_packages_dir(self) -> str:
        site_packages = Path.home() / f".local/lib/python{self.python_major_minor}/site-packages"
        site_packages.mkdir(parents=True, exist_ok=True)
        return str(site_packages)

    def _get_flatpak_site_packages_dir(self) -> str:
        sand_box = os.getenv("FLATPAK_SANDBOX_DIR", "") or str(
            Path.home() / ".var/app/org.libreoffice.LibreOffice/sandbox"
        )
        site_packages = Path(sand_box) / f"lib/python{self.python_major_minor}/site-packages"
        site_packages.mkdir(parents=True, exist_ok=True)
        return str(site_packages)

    def _get_mac_site_packages_dir(self) -> str:
        # sourcery skip: class-extract-method
        site_packages = Path.home() / f"Library/LibreOfficePython/{self.python_major_minor}/lib/python/site-packages"
        site_packages.mkdir(parents=True, exist_ok=True)
        return str(site_packages)

    @property
    def url_pip(self) -> str:
        return self._url_pip

    """
    String path such as ``https://bootstrap.pypa.io/get-pip.py``

    The value for this property can be set in pyproject.toml (tool.oxt.token.url_pip)
    """

    @property
    def python_path(self) -> Path:
        """
        Gets the path to the python executable.

        For some strange reason, on windows, the path can come back as 'soffice.bin' for 'sys.executable'.
        """
        return self._python_path

    @property
    def log_file(self) -> str:
        """
        Gets the name of the log file.

        The value for this property can be set in pyproject.toml (tool.oxt.token.log_file)
        """
        return self._log_file

    @property
    def log_name(self) -> str:
        """
        Gets the name of the log file.

        The value for this property can be set in pyproject.toml (tool.oxt.token.log_name)
        """
        return self._log_name

    @property
    def log_level(self) -> int:
        """
        Gets the log level.

        The value for this property can be set in pyproject.toml (tool.oxt.token.log_level)
        """
        return self._log_level

    @property
    def log_format(self) -> str:
        """
        Gets the log format.

        The value for this property can be set in pyproject.toml (tool.oxt.token.log_format)
        """
        return self._log_format

    @property
    def py_pkg_dir(self) -> str:
        """
        Gets the name of the directory where python packages are installed as a zip.

        The value for this property can be set in pyproject.toml (tool.oxt.token.py_pkg_dir)
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
    def is_win(self) -> bool:
        """
        Gets the flag indicating if the operating system is Windows.
        """
        return self._is_win

    @property
    def is_mac(self) -> bool:
        """
        Gets the flag indicating if the operating system is macOS.
        """
        return self._is_mac

    @property
    def is_app_image(self) -> bool:
        """
        Gets the flag indicating if LibreOffice is running as an AppImage.
        """
        return self._is_app_image

    @property
    def is_flatpak(self) -> bool:
        """
        Gets the flag indicating if LibreOffice is running as a Flatpak.
        """
        return self._is_flatpak

    @property
    def os(self) -> str:
        """
        Gets the operating system.
        """
        return self._os

    @property
    def pip_wheel_url(self) -> str:
        """
        Gets the pip wheel url.

        May be empty string.
        """
        return self._pip_wheel_url

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
    def python_major_minor(self) -> str:
        """
        Gets the python major minor version, such as ``3.9``
        """
        return self._python_major_minor

    @property
    def site_packages(self) -> str:
        """
        Gets the path to the site-packages directory. May be empty string.

        Is valid for ``Flatpak``, ``MAC``, ``App Image``, ``Linux``.
        """
        return self._site_packages
