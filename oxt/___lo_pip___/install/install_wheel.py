"""After pip is installed this module can install Wheel."""
from __future__ import annotations
from ..config import Config


class InstallWheel:
    """
    Installs Wheel package if it is not already installed.
    """

    def __init__(self) -> None:
        self._config = Config()

    def install(self) -> bool:
        """
        Install wheel if it is not already installed.

        Returns:
            bool: True if successful, False otherwise.
        """
        if self._config.is_flatpak:
            return self._install_flatpak()

        return self._install_default()

    def _install_flatpak(self) -> bool:
        from .pkg_installers.install_pkg_flatpak import InstallPkgFlatpak

        installer = InstallPkgFlatpak()
        return installer.install({"wheel": ""})

    def _install_default(self) -> bool:
        from .pkg_installers.install_pkg import InstallPkg

        installer = InstallPkg()
        return installer.install({"wheel": ""})
