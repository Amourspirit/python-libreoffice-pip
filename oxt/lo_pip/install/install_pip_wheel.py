from __future__ import annotations
import tempfile
from pathlib import Path

from ..config import Config
from .download import Download
from ..oxt_logger import OxtLogger
from ..input_output import file_util


class InstallPipWheel:
    """Download and install PIP from wheel url"""

    def __init__(self) -> None:
        self._config = Config()
        self._logger = OxtLogger(log_name=__name__)

    def install_pip_wheel(self, dst: str | Path = "") -> None:
        """
        Installs a pip wheel file.

        Args:
            dst (str | Path, Optional): The destination directory where the pip wheel file will be installed. If not provided, the ``pythonpath`` location will be used.

        Returns:
            None

        Raises:
            None
        """
        url = self._config.pip_wheel_url
        if not url:
            self._logger.error("PIP installation has failed - No wheel url")
            return

        if not dst:
            root_pth = Path(file_util.get_package_location(self._config.lo_identifier))
            dst = root_pth / "pythonpath"

        with tempfile.TemporaryDirectory() as temp_dir:
            # temp_dir = tempfile.gettempdir()
            path_pip = Path(temp_dir)

            filename = path_pip / "pip-wheel.whl"

            dl = Download()
            data, _, err = dl.url_open(url, verify=False)
            if err:
                self._logger.error("Unable to download PIP installation wheel file")
                return
            dl.save_binary(pth=filename, data=data)

            if filename.exists():
                self._logger.info("PIP wheel file has been saved")
            else:
                self._logger.error("PIP wheel file has not been saved")
                return

            self._unzip_wheel(filename=filename, dst=dst)

    def _unzip_wheel(self, filename: Path, dst: str | Path) -> None:
        """Unzip the downloaded wheel file"""
        if isinstance(dst, str):
            dst = Path(dst)
        try:
            file_util.unzip_file(filename, dst)
            if dst.exists():
                self._logger.debug(f"PIP wheel file has been unzipped to {dst}")
            else:
                self._logger.error("PIP wheel file has not been unzipped")
                return
        except Exception as err:
            self._logger.error("Unable to unzip wheel file", exc_info=True)
            return
