from __future__ import annotations
import subprocess
import tempfile
from pathlib import Path
from typing import List
import urllib.request
import os.path
import logging
import uno

from com.sun.star.uno import RuntimeException

from ..config import Config
from ..input_output import file_util


logger = logging.getLogger(Config().log_name)
formatter = logging.Formatter(Config().log_format)

if Config().log_level <= 0:
    log_handler = logging.StreamHandler()
    log_handler.setFormatter(formatter)
else:
    try:
        log_handler = logging.FileHandler(
            os.path.join(uno.fileUrlToSystemPath(file_util.get_user_profile_path()), Config().log_file),
            mode="w",
            delay=True,
        )
        log_handler.setFormatter(formatter)
    except RuntimeException:
        # At installation time, no context is available -> just ignore it.
        pass


class InstallPip:
    """Singleton class for the PIP install."""

    def __init__(self) -> None:
        self._config = Config()
        self.path_python = Path(self._config.python_path)

    def install_pip(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            # Do something with the temporary directory
            # print(f"Temporary directory created at {temp_dir}")
            path_pip = Path(temp_dir)

            url = self._config.url_pip
            filename = path_pip / "get-pip.py"

            urllib.request.urlretrieve(url, filename)

            if not filename.exists():
                # msg = ("Unable to copy PIP installation file",)

                return

            # PIP installation file has been saved

            try:
                # "Starting PIP installation…"
                logger.info("Starting PIP installation…")
                cmd = [str(self.path_python), f"{path_pip}", "--user"]

                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                stdout, stderr = process.communicate()
                str_stderr = stderr.decode("utf-8")
                if process.returncode != 0:
                    # "PIP installation has failed, see log"
                    logger.error("PIP installation has failed")
                    if str_stderr:
                        logger.error(str_stderr)
                    return

                cmd = self._cmd_pip("--version")
                process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                if process.returncode == 0:
                    # "PIP was installed successfully"
                    logger.info("PIP was installed successfully")
                else:
                    # "PIP installation has failed, see log"
                    logger.error("PIP installation has failed")
            except Exception as err:
                # "PIP installation has failed, see log"
                logger.error("PIP installation has failed")
                logger.error(err)

            return

    def install_requirements(self, fnm: str | Path) -> None:
        """Install the requirements."""
        logger.info("install_requirements - Installing requirements…")
        if not self.is_pip_installed():
            self.install_pip()
        if not self.is_pip_installed():
            logger.error("install_requirements - PIP installation has failed")
            return
        self._install(path=str(fnm))
        return

    def install_package(self, value: str) -> None:
        """Install a package."""
        if not self.is_pip_installed():
            self.install_pip()
        if not self.is_pip_installed():
            return
        self._install(value=value)
        return

    def pip_upgrade(self) -> None:
        """Upgrade PIP."""
        logger.info("pip_upgrade - Upgrading PIP…")
        if not self.is_pip_installed():
            self.install_pip()
        if not self.is_pip_installed():
            return
        self._install()
        return

    def _cmd_pip(self, *args: str) -> List[str]:
        cmd: List[str] = [str(self.path_python), "-m", "pip", *args]
        return cmd

    def _install(self, value: str = "", path: str = ""):
        cmd = ["install", "--upgrade", "--user"]
        msg = "Install - Upgrading pip success!"
        err_msg = "Install - Upgrading pip failed!"
        if value:
            name = value.split()[0].strip()
            cmd = self._cmd_pip(*[*cmd, *name])
            msg = f"Install - Installing {name} success!"
            err_msg = f"Install - Installing {name} failed!"
        else:
            cmd = self._cmd_pip(*[*cmd, "-r", f"{path}"])
            msg = "Install - Installing requirements success!"
            err_msg = "Install - Installing requirements failed!"
        process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if process.returncode == 0:
            logger.info(msg)
        else:
            logger.error(err_msg)
        return

    def is_pip_installed(self) -> bool:
        """Check if PIP is installed."""
        cmd = self._cmd_pip("--version")
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode == 0:
            return True
        return False

    # def cmd_shell_action(self, event: Any) -> None:
    #     cmd = None
    #     if app.IS_WIN:
    #         cmd = [str(self.path_python)]
    #     else:
    #         cmd = ["exec"]  # 'exec "{}"'
    #         if app.IS_MAC:
    #             cmd = ["open"]  # 'open "{}"'
    #         elif app.DESKTOP == "gnome":
    #             cmd = ["gnome-terminal", "--"]  # "gnome-terminal -- {}"

    #         cmd = cmd.append(str(self.path_python))
    #     if cmd:
    #         subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    #     return
