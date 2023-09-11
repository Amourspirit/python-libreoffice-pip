import sys
import os
import subprocess

# import pkg_resources
from importlib.metadata import distribution, PackageNotFoundError
from pathlib import Path

import logging
import uno

from com.sun.star.uno import RuntimeException
from ..config import Config
from ..input_output import file_util


logger = logging.getLogger(Config().log_name)
formatter = logging.Formatter(Config().log_format)

if Config().log_level <= 0:
    log_handler = logging.StreamHandler()
    log_handler.setFormatter(formatter)
else:
    try:
        log_handler = logging.FileHandler(
            os.path.join(uno.fileUrlToSystemPath(file_util.get_user_profile_path()), Config().log_file),
            mode="w",
            delay=True,
        )
        log_handler.setFormatter(formatter)
    except RuntimeException:
        # At installation time, no context is available -> just ignore it.
        pass

# https://docs.python.org/3.8/library/importlib.metadata.html#module-importlib.metadata
# https://stackoverflow.com/questions/44210656/how-to-check-if-a-module-is-installed-in-python-and-if-not-install-it-within-t


class InstallPkg:
    """Install pip packages."""

    def __init__(self) -> None:
        self._config = Config()
        self.path_python = Path(self._config.python_path)
