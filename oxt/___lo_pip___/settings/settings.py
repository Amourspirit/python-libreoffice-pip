from __future__ import annotations
from typing import Any, Dict, cast, TYPE_CHECKING
import uno

from ..lo_util.configuration import Configuration
from ..meta.singleton import Singleton
from ..oxt_logger import OxtLogger
from ..config import Config

from ..events.lo_events import Events
from ..events.args import EventArgs
from ..events.named_events import ConfigurationNamedEvent

if TYPE_CHECKING:
    from com.sun.star.configuration import ConfigurationAccess


class Settings(metaclass=Singleton):
    """Singleton Class. Manages Settings for the extension."""

    def __init__(self) -> None:
        self._logger = OxtLogger(log_name=__name__)
        self._configuration = Configuration()
        cfg = Config()
        self._lo_identifier = cfg.lo_identifier
        self._lo_implementation_name = cfg.lo_implementation_name
        self._current_settings = self.get_settings()
        self._events = Events(source=self)
        self._set_events()

    def _set_events(self) -> None:
        def on_configuration_saved(src: Any, event_args: EventArgs) -> None:
            self._logger.debug("Settings. Configuration saved. Updating settings..")
            self._update_settings()

        def on_configuration_str_lst_saved(src: Any, event_args: EventArgs) -> None:
            self._logger.debug("Settings. Configuration str lst saved. Updating settings..")
            self._update_settings()

        # keep callbacks in scope
        self._fn_on_configuration_saved = on_configuration_saved
        self._fn_oon_configuration_str_lst_saved = on_configuration_str_lst_saved

        events = self._events  # _Events()

        events.on(event_name=ConfigurationNamedEvent.CONFIGURATION_SAVED, callback=on_configuration_saved)
        events.on(
            event_name=ConfigurationNamedEvent.CONFIGURATION_STR_LST_SAVED, callback=on_configuration_str_lst_saved
        )

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

    def _update_settings(self) -> None:
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
