"""After pip is installed this module can install Wheel."""
from __future__ import annotations
from ..config import Config


class InstallWheel:
    """
    Installs Wheel package if it is not already installed.
    """

    def __init__(self) -> None:
        self._config = Config()

    def install(self) -> None:
        """Install wheel if it is not already installed."""
        if self._config.is_flatpak:
            self._install_flatpak()
        else:
            self._install_default()

    def _install_flatpak(self) -> None:
        from .pkg_installers.install_pkg_flatpak import InstallPkgFlatpak

        installer = InstallPkgFlatpak()
        installer.install({"wheel": ""})

    def _install_default(self) -> None:
        from .pkg_installers.install_pkg import InstallPkg

        installer = InstallPkg()
        installer.install({"wheel": ""})
