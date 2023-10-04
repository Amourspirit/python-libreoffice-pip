from __future__ import annotations
from typing import cast, Dict
from pathlib import Path
import json
import toml

from ..meta.singleton import Singleton
from ..config import Config
from .token import Token


class JsonConfig(metaclass=Singleton):
    """Singleton Class the Config Json."""

    def __init__(self) -> None:
        token = Token()
        self._config = Config()
        cfg = toml.load(self._config.toml_path)
        self._requirements = cast(Dict[str, str], cfg["tool"]["oxt"]["requirements"])
        try:
            self._zip_preinstall_pure = bool(cfg["tool"]["oxt"]["config"]["zip_preinstall_pure"])
        except Exception:
            self._zip_preinstall_pure = False
        try:
            self._auto_install_in_site_packages = cast(
                bool, cfg["tool"]["oxt"]["config"]["auto_install_in_site_packages"]
            )
        except Exception:
            self._auto_install_in_site_packages = False
        try:
            self._install_wheel = cast(bool, cfg["tool"]["oxt"]["config"]["install_wheel"])
        except Exception:
            self._install_wheel = False
        # try:
        #     self._log_pip_installs = cast(bool, cfg["tool"]["oxt"]["config"]["log_pip_installs"])
        # except Exception:
        #     self._log_pip_installs = True

    def update_json_config(self, json_config_path: Path) -> None:
        """Read and updates the config.json file."""
        with open(json_config_path, "r") as f:
            json_config = json.load(f)
        token = Token()
        json_config["py_pkg_dir"] = token.get_token_value("py_pkg_dir")
        json_config["lo_identifier"] = token.get_token_value("lo_identifier")
        json_config["lo_implementation_name"] = token.get_token_value("lo_implementation_name")

        json_config["zipped_preinstall_pure"] = self._zip_preinstall_pure
        json_config["auto_install_in_site_packages"] = self._auto_install_in_site_packages
        json_config["install_wheel"] = self._install_wheel
        # json_config["log_pip_installs"] = self._log_pip_installs
        # update the requirements
        json_config["requirements"] = self._requirements
        json_config["has_locals"] = self._config.has_locals

        # save the file
        with open(json_config_path, "w") as f:
            json.dump(json_config, f, indent=4)
