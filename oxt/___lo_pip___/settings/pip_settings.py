from __future__ import annotations
from typing import cast, Tuple

from .settings import Settings
from ..lo_util.configuration import Configuration
from ..config import Config


class PipInfoSettings(Settings):
    """Singleton Class. Manages Settings for the extension."""

    def __init__(self) -> None:
        super().__init__()
        self._config = Config()
        self._installed_local_pips = cast(Tuple, self.current_settings.get("InstalledLocalPips", ()))
        self._node_value = self._get_node_value()

    def _get_node_value(self) -> str:
        return f"/{self.lo_implementation_name}.Settings/PipInfo"

    def append_installed_local_pip(self, pip_name: str) -> None:
        """Appends a pip to the installed local pips."""
        if pip_name not in self.installed_local_pips:
            self.installed_local_pips = (*self.installed_local_pips, pip_name)

    def remove_installed_local_pip(self, pip_name: str) -> None:
        """Removes a pip from the installed local pips."""
        if pip_name in self.installed_local_pips:
            self.installed_local_pips = tuple(pip for pip in self.installed_local_pips if pip != pip_name)

    @property
    def installed_local_pips(self) -> Tuple[str, ...]:
        """Gets/Sets the installed local pips."""
        return self._installed_local_pips

    @installed_local_pips.setter
    def installed_local_pips(self, value: Tuple[str, ...]) -> None:
        self.configuration.save_configuration_str_lst(
            node_value=self._node_value, name="InstalledLocalPips", value=value
        )
        self._installed_local_pips = value
        self.on_update_settings()
