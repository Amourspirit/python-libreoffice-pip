from __future__ import annotations
from typing import cast, Any, Dict, List
import toml
from ..meta.singleton import Singleton
from .. import file_util


class Token(metaclass=Singleton):
    """Singleton Class the tokens."""

    def __init__(self) -> None:
        toml_path = file_util.find_file_in_parent_dirs("pyproject.toml")
        cfg = toml.load(toml_path)

        tokens = cast(Dict[str, str], cfg["tool"]["oxt"]["token"])
        self._tokens: Dict[str, str] = {}
        for token, replacement in tokens.items():
            self._tokens[f"___{token}___"] = replacement
        self._tokens["___version___"] = str(cfg["tool"]["poetry"]["version"])
        self._tokens["___license___"] = str(cfg["tool"]["poetry"]["license"])
        self._tokens["___oxt_name___"] = str(cfg["tool"]["oxt"]["config"]["oxt_name"])
        self._tokens["___dist_dir___"] = str(cfg["tool"]["oxt"]["config"]["dist_dir"])
        self._tokens["___update_file___"] = str(cfg["tool"]["oxt"]["config"]["update_file"])
        self._tokens["___py_pkg_dir___"] = str(cfg["tool"]["oxt"]["config"]["py_pkg_dir"])
        self._tokens["___authors___"] = self._get_authors(cfg)
        for token, replacement in self._tokens.items():
            self._tokens[token] = self.process(replacement)
        self._validate()

    # region Methods
    def _validate(self) -> None:
        keys = {
            "lo_identifier",
            "lo_implementation_name",
            "display_name_en_us",
            "description_en_us",
            "publisher_en_us",
            "url_pip",
            "log_format",
            "lo_pip",
            "platform",
        }
        for key in keys:
            str_key = f"___{key}___"
            if str_key not in self._tokens:
                raise ValueError(f"Token '{key}' not found")
            if not self._tokens[str_key]:
                raise ValueError(f"Token '{key}' is empty")
        if "___log_level___" not in self._tokens:
            raise ValueError("Token 'log_level' not found")
        levels = {"none", "debug", "info", "warning", "error", "critical"}
        if self._tokens["___log_level___"].lower() not in levels:
            raise ValueError(f"Token 'log_level' is invalid: {self._tokens['___log_level___']}")
        lo_identifier = self.get_token_value("lo_identifier")
        if lo_identifier != "org.openoffice.extensions.ooopip":
            if self.get_token_value("lo_implementation_name") == "OooPipRunner":
                raise ValueError(
                    "Token 'lo_implementation_name' value is invalid and must be renamed in tool.oxt.token in pyproject.toml. Every project must have unique lo_implementation_name value."
                )
            if self.get_token_value("lo_pip") == "lo_pip":
                raise ValueError(
                    "Token 'lo_pip' value is invalid and must be renamed in tool.oxt.token in pyproject.toml. Every project must have unique lo_pip value."
                )
            log_file = self.get_token_value("log_file")
            if log_file == "pip_install.log":
                raise ValueError(
                    "Token 'log_file' value is invalid and must be renamed in tool.oxt.token in pyproject.toml. Every project must have unique log_file value or set to empty string for no logging."
                )

    def process(self, text: str) -> str:
        """Processes the given text."""
        for token, replacement in self._tokens.items():
            text = text.replace(token, replacement)

        return text

    def get_token_value(self, token: str) -> str:
        """
        Returns the value of the given token.

        Args:
            token (str): The token in normal form (without '___' at the beginning and end).
        """
        return self._tokens.get(f"___{token}___", "")

    def _get_authors(self, cfg: Dict[str, Any]) -> str:
        """Returns the authors."""
        authors = cast(List[str], cfg["tool"]["poetry"]["authors"])
        results: List[str] = []
        for author in authors:
            results.append(author.split("<")[0].strip())
        return ", ".join(results)

    # endregion Methods

    # region Properties
    @property
    def tokens(self) -> Dict[str, str]:
        """The tokens."""
        return self._tokens

    # endregion Properties
