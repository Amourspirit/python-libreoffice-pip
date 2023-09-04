# coding: utf-8
from __future__ import annotations
from pathlib import Path
import json
import sys
import logging

logger = logging.getLogger(__name__)
logger.debug("config.py imported")

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
            else:
                # provide defaults because at this time scriptmerge
                # does not include non *.py files when it packages scripts
                data = {
                    "url_pip": "https://bootstrap.pypa.io/get-pip.py",
                    "log_file": "lo_pip.log",
                    "log_name": "OOO PIP Installer",
                    "": "INFO",
                    "log_format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
                    "py_pkg_dir": "py_pkgs",
                }

            cls._instance = super().__call__(**data)
        return cls._instance


class Config(metaclass=ConfigMeta):
    """
    Singleton Configuration Class

    Generally speaking this class is only used internally.
    """

    def __init__(self, **kwargs):
        logger.debug("Config.__init__ called")
        self._url_pip = str(kwargs["url_pip"])
        self._log_file = str(kwargs["log_file"])
        self._log_name = str(kwargs["log_name"])
        self._log_format = str(kwargs["log_format"])
        self._py_pkg_dir = str(kwargs["py_pkg_dir"])
        python_path = file_util.get_which("python")
        if not python_path:
            python_path = sys.executable
        if not python_path:
            raise FileNotFoundError("python not found")
        self._python_path = Path(python_path)
        self._log_level = self._get_log_level(kwargs["log_level"])
        logger.debug("Config.__init__ completed")

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
