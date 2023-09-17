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
        self._config = Config()
        cfg = toml.load(self._config.toml_path)
        self._requirements = cast(Dict[str, str], cfg["tool"]["oxt"]["requirements"])

    def update_json_config(self, json_config_path: Path) -> None:
        """Read and updates the config.json file."""
        with open(json_config_path, "r") as f:
            json_config = json.load(f)
        token = Token()
        json_config["url_pip"] = token.tokens["___url_pip___"]
        json_config["log_name"] = token.tokens["___log_name___"]
        json_config["log_level"] = token.tokens["___log_level___"]
        json_config["log_file"] = token.tokens["___log_file___"]
        json_config["log_format"] = token.tokens["___log_format___"]
        json_config["py_pkg_dir"] = token.tokens["___py_pkg_dir___"]

        # update the requirements
        json_config["requirements"] = self._requirements

        # save the file
        with open(json_config_path, "w") as f:
            json.dump(json_config, f, indent=4)
