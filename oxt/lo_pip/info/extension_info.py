from __future__ import annotations
from typing import cast, Any, Tuple
import uno

from com.sun.star.deployment import XPackageInformationProvider

from ..meta.singleton import Singleton
from ..oxt_logger import OxtLogger


class ExtensionInfo(metaclass=Singleton):
    """
    Gets info for an installed extension in LibreOffice.

    Singleton Class.
    """

    def __init__(self) -> None:
        self._logger = OxtLogger(log_name=__name__)

    def get_extension_info(self, id: str) -> Tuple[str, ...]:
        """
        Gets info for an installed extension in LibreOffice.

        Args:
            id (str): Extension id

        Returns:
            Tuple[str, ...]: Extension info
        """
        try:
            pip = self.get_pip()
        except Exception:
            return ()
        exts_tbl = pip.getExtensionList()
        for el in exts_tbl:
            if el[0] == id:
                return el
        self._logger.info(f"Extension {id} is not found")
        return ()

    def get_pip(self) -> XPackageInformationProvider:
        """
        Gets Package Information Provider

        Raises:
            Exception: if unable to obtain XPackageInformationProvider interface

        Returns:
            XPackageInformationProvider: Package Information Provider
        """
        try:
            ctx: Any = uno.getComponentContext()
            pip = ctx.getValueByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
            if pip is None:
                raise Exception("Unable to get PackageInformationProvider, pip is None")
            return cast(XPackageInformationProvider, pip)
        except Exception as err:
            self._logger.error(err, exc_info=True)
            raise

    def get_extension_loc(self, id: str) -> str:
        """
        Gets location for an installed extension in LibreOffice

        Args:
            id (str): Extension id

        Returns:
            str: Extension location on success; Otherwise, an empty string
        """
        try:
            pip = self.get_pip()
        except Exception:
            return ""
        return pip.getPackageLocation(id)

    def log_extensions(self) -> None:
        """
        Logs extensions to log file
        """
        try:
            pip = self.get_pip()
        except Exception:
            self._logger.info("No package info provider found")
            return
        exts_tbl = pip.getExtensionList()
        self._logger.info("Extensions:")
        for i in range(len(exts_tbl)):
            self._logger.info(f"{i+1}. ID: {exts_tbl[i][0]}")
            self._logger.info(f"   Version: {exts_tbl[i][1]}")
            self._logger.info(f"   Loc: {pip.getPackageLocation(exts_tbl[i][0])}")
