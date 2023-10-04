from __future__ import annotations
from typing import cast, Tuple

from .settings import Settings
from ..meta.singleton import Singleton
from ..lo_util.configuration import Configuration
from ..config import Config


class GeneralSettings(metaclass=Singleton):
    """Singleton Class. Manages Settings for the extension."""

    def __init__(self) -> None:
        settings = Settings()
        self._configuration = Configuration()
        self._node_value = f"/{settings.lo_implementation_name}.Settings/GeneralSettings"
        self._lo_implementation_name = str(settings.current_settings.get("LoImplementationName", ""))
        self._lo_identifier = str(settings.current_settings.get("LoIdentifier", ""))
        self._publisher_url = str(settings.current_settings.get("PublisherUrl", ""))
        self._update_url_xml = str(settings.current_settings.get("UpdateUrlXml", ""))
        self._update_url_oxt = str(settings.current_settings.get("UpdateUrlOxt", ""))
        self._url_pip = str(settings.current_settings.get("UrlPip", ""))
        self._pip_wheel_url = str(settings.current_settings.get("UrlPipWheel", ""))
        self._test_internet_url = str(settings.current_settings.get("UrlTestInternet", ""))
        self._platform = str(settings.current_settings.get("Platform", ""))
        self._log_pip_installs = bool(settings.current_settings.get("LogPipInstalls", False))
        self._show_progress = bool(settings.current_settings.get("ShowProgress", False))

    @property
    def log_pip_installs(self) -> bool:
        """
        Gets the flag indicating if pip installs should be logged.
        """
        return self._log_pip_installs

    @property
    def lo_implementation_name(self) -> str:
        """Gets the name of the LibreOffice extension implementation."""
        return self._lo_implementation_name

    @property
    def lo_identifier(self) -> str:
        """Gets the identifier of the LibreOffice extension."""
        return self._lo_identifier

    @property
    def publisher_url(self) -> str:
        """Gets the publisher url of the LibreOffice extension."""
        return self._publisher_url

    @property
    def update_url_xml(self) -> str:
        """Gets the XML update url of the LibreOffice extension."""
        return self._update_url_xml

    @property
    def url_pip(self) -> str:
        """Gets the url to ``get-pip.py`` that installs pip."""
        return self._url_pip

    @property
    def pip_wheel_url(self) -> str:
        """Gets the url to the pip wheel."""
        return self._pip_wheel_url

    @property
    def platform(self) -> str:
        """Gets the platform of the LibreOffice extension is targeted for."""
        return self._platform

    @property
    def show_progress(self) -> bool:
        """Gets the flag indicating if the terminal should be shown."""
        return self._show_progress

    @property
    def test_internet_url(self) -> str:
        """
        String path such as ``https://www.google.com``

        The value for this property can be set in pyproject.toml (tool.oxt.token.test_internet_url)
        """
        return self._test_internet_url
