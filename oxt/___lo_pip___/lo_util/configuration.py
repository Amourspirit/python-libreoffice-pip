from __future__ import annotations
from typing import Any, cast, Dict, Tuple, TYPE_CHECKING
from typing import TypedDict
import uno
from com.sun.star.beans import PropertyValue

from .util import Util
from ..meta.singleton import Singleton

if TYPE_CHECKING:
    from com.sun.star.configuration import ConfigurationUpdateAccess  # service

    class SettingsT(TypedDict):
        # https://peps.python.org/pep-0589/
        names: Tuple[str, ...]
        """Names of the properties."""
        values: Tuple[Any, ...]
        """Values of the properties."""

else:
    SettingsT = object


class Configuration(metaclass=Singleton):
    """Configuration class for accessing and saving configuration settings stored in LibreOffice."""

    def get_configuration_access(self, node_value: str, updatable: bool = False) -> Any:
        """
        Access configuration value.

        Args:
            node_value (str): The configuration key node as a string.
            updatable (bool, optional): Set True when accessor needs to modify the key value. Defaults to False.

        Returns:
            Any: The configuration value.
        """
        util = Util()
        cp = util.create_uno_service("com.sun.star.configuration.ConfigurationProvider")
        node = PropertyValue("nodepath", 0, node_value, 0)
        if updatable:
            return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
        else:
            return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))

    def save_configuration(self, node_value: str, settings: SettingsT) -> None:
        """
        Save Configuration settings.

        Args:
            node_value (str): Node value such as "/OooPipRunner.Settings/Help"
            settings (dict): A dictionary with the following keys:
                - names: Tuple[str, ...]
                - values: Tuple[Any, ...]

        Raises:
            Exception: If the configuration cannot be saved.

        Returns:
            None: None


        See Also:
            ``Configuration.convert_dict_to_settings()``
        """
        if not settings:
            return
        try:
            writer = cast("ConfigurationUpdateAccess", self.get_configuration_access(node_value, True))
            writer.setPropertyValues(settings["names"], settings["values"])
            # uno.invoke(writer, "setPropertyValue", (settings["names"], settings["values"]))  # type: ignore
            writer.commitChanges()

        except Exception as err:
            # self._logger.error(f"Error saving configuration: {err}", exc_info=True)
            raise err

    def save_configuration_str_lst(self, node_value: str, name: str, value: Tuple[str, ...]) -> None:
        # https://forum.openoffice.org/en/forum/viewtopic.php?t=56460
        try:
            vals = uno.Any("[]string", value)  # type: ignore
            writer = cast("ConfigurationUpdateAccess", self.get_configuration_access(node_value, True))
            uno.invoke(writer, "replaceByName", (name, vals))  # type: ignore
            writer.commitChanges()

        except Exception as err:
            # self._logger.error(f"Error saving configuration: {err}", exc_info=True)
            raise err

    def convert_dict_to_settings(self, input_dict: Dict[str, Any]) -> SettingsT:
        """
        Convert a dictionary to a SettingsT for use with ``save_configuration()``.

        The ``input_dict`` can contain any number of keys and values. The keys and values are converted to tuples.
        The keys are stored in the ``names`` tuple and the values are stored in the ``values`` tuple.

        All keys must be of type ``str``.

        Args:
            settings (input_dict): A dictionary with the following keys:
                - names: Tuple[str, ...]
                - values: Tuple[Any, ...]

        Returns:
            SettingsT: A SettingsT object.
        """
        if not input_dict:
            return cast(SettingsT, {})
        tuple_keys = tuple(str(key) for key in input_dict.keys())
        tuple_values = tuple(input_dict.values())
        return cast(SettingsT, {"names": tuple_keys, "values": tuple_values})
