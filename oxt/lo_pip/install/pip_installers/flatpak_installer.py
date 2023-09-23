from __future__ import annotations
import sys

from ...config import Config
from ...oxt_logger import OxtLogger

from .base_installer import BaseInstaller


class FlatpakInstaller(BaseInstaller):
    """Class for the Flatpak PIP install."""

    def _get_logger(self) -> OxtLogger:
        return OxtLogger(log_name=__name__)

    def install_pip(self) -> None:
        if self.is_pip_installed():
            self._logger.info("PIP is already installed")
            return
        if self._install_wheel():
            if self.is_pip_installed():
                self._logger.info("PIP was installed successfully")
        else:
            self._logger.error("PIP installation has failed")

        return

    def _install_wheel(self) -> bool:
        cfg = Config()
        result = False
        from ..install_pip_wheel import InstallPipWheel

        installer = InstallPipWheel()
        try:
            installer.install_pip_wheel(cfg.site_packages)
            if cfg.site_packages not in sys.path:
                sys.path.append(cfg.site_packages)
            result = True
        except Exception as err:
            self._logger.error(err)
            return result
        return result
