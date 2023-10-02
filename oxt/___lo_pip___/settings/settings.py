from __future__ import annotations
from typing import Any, Dict, cast, TYPE_CHECKING
import uno

from ..lo_util.configuration import Configuration
from ..meta.singleton import Singleton
from ..oxt_logger import OxtLogger
from ..config import Config

if TYPE_CHECKING:
    from com.sun.star.configuration import ConfigurationAccess


class Settings(metaclass=Singleton):
    """Manages Settings for the extension."""

    def __init__(self) -> None:
        self._logger = OxtLogger(log_name=__name__)
        self._configuration = Configuration()
        cfg = Config()
        self._lo_identifier = cfg.lo_identifier
        self._lo_implementation_name = cfg.lo_implementation_name
        self._current_settings = self.get_settings()

    def _get_node_value(self) -> str:
        raise NotImplementedError

    def get_settings(self) -> Dict[str, Any]:
        key = f"/{self.lo_implementation_name}.Settings"
        reader = cast("ConfigurationAccess", self._configuration.get_configuration_access(key))
        group_names = reader.getElementNames()
        settings = {}
        for groupname in group_names:
            group = cast("ConfigurationAccess", reader.getByName(groupname))
            props = group.getElementNames()
            values = group.getPropertyValues(props)
            settings.update({k: v for k, v in zip(props, values)})
        self._logger.debug(f"Returning {self.lo_implementation_name} settings.")
        return settings

    def on_update_settings(self) -> None:
        """Updates the current settings"""
        self._current_settings = self.get_settings()

    @property
    def lo_identifier(self) -> str:
        return self._lo_identifier

    @property
    def lo_implementation_name(self) -> str:
        return self._lo_implementation_name

    @property
    def current_settings(self) -> Dict[str, Any]:
        return self._current_settings

    @property
    def configuration(self) -> Configuration:
        return self._configuration
