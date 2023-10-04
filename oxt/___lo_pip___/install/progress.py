from __future__ import annotations
from typing import TYPE_CHECKING
import subprocess
from ..input_output.proc import kill_proc

if TYPE_CHECKING:
    from ..config import Config


class Progress:
    def __init__(self, start_msg: str):
        if not TYPE_CHECKING:
            from ..config import Config
        self._start_msg = start_msg
        self._config = Config()
        self._proc = None
        self._pid = -1

    def _get_code(self) -> str:
        code = f"""
import time
print("{self._start_msg} ", flush=True, end="")
while True:
    print(".", flush=True, end="")
    time.sleep(1)
"""
        return code

    def start(self) -> None:
        """Start the progress indicator as a terminal window."""
        self._proc = subprocess.Popen(
            [str(self._config.python_path), "-c", self._get_code()],
            creationflags=subprocess.CREATE_NEW_CONSOLE | subprocess.CREATE_NEW_PROCESS_GROUP,
        )
        self._pid = self._proc.pid

    def kill(self) -> None:
        if self._proc != -1:
            kill_proc(self._pid)
            self._proc = None
