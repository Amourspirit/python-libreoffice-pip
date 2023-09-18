from __future__ import annotations
import subprocess
import tempfile
from pathlib import Path
from typing import List
import urllib.request

from ..config import Config
from ..oxt_logger import OxtLogger


# https://stackoverflow.com/search?q=%5Bpython%5D+run+subprocess+without+popup+terminal
# silent subprocess
_si = subprocess.STARTUPINFO()
_si.dwFlags |= subprocess.STARTF_USESHOWWINDOW


class InstallPip:
    """Singleton class for the PIP install."""

    def __init__(self) -> None:
        self._config = Config()
        self.path_python = Path(self._config.python_path)
        self._logger = OxtLogger(log_name=__name__)

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
                self._logger.error("Unable to copy PIP installation file")
                return

            # PIP installation file has been saved

            try:
                # "Starting PIP installation…"
                self._logger.info("Starting PIP installation…")
                cmd = [str(self.path_python), f"{filename}", "--user"]

                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=_si)
                stdout, stderr = process.communicate()
                str_stderr = stderr.decode("utf-8")
                if process.returncode != 0:
                    # "PIP installation has failed, see log"
                    self._logger.error("PIP installation has failed")
                    if str_stderr:
                        self._logger.error(str_stderr)
                    return

                cmd = self._cmd_pip("--version")
                process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=_si)
                if process.returncode == 0:
                    # "PIP was installed successfully"
                    self._logger.info("PIP was installed successfully")
                else:
                    # "PIP installation has failed, see log"
                    self._logger.error("PIP installation has failed")
            except Exception as err:
                # "PIP installation has failed, see log"
                self._logger.error("PIP installation has failed")
                self._logger.error(err)

            return

    def install_requirements(self, fnm: str | Path) -> None:
        """Install the requirements."""
        self._logger.info("install_requirements - Installing requirements…")
        if not self.is_pip_installed():
            self.install_pip()
        if not self.is_pip_installed():
            self._logger.error("install_requirements - PIP installation has failed")
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
        self._logger.info("pip_upgrade - Upgrading PIP…")
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
        process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=_si)
        if process.returncode == 0:
            self._logger.info(msg)
        else:
            self._logger.error(err_msg)
        return

    def is_pip_installed(self) -> bool:
        """Check if PIP is installed."""
        cmd = self._cmd_pip("--version")
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=_si)
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
