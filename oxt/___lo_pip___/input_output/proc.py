from __future__ import annotations
import sys
import subprocess

_IS_WINDOWS_PLATFORM = "win" in sys.platform


def kill_proc(pid: int) -> None:
    """Kill a process by pid."""
    if _IS_WINDOWS_PLATFORM:
        _kill_windows(pid)
    else:
        _kill_linux(pid)


def _kill_windows(pid: int) -> None:
    """Kill a process by pid on Windows."""
    subprocess.call(["taskkill", "/F", "/T", "/PID", str(pid)])


def _kill_linux(pid: int) -> None:
    """Kill a process by pid on Linux."""
    # https://alexandra-zaharia.github.io/posts/kill-subprocess-and-its-children-on-timeout-python/
    os.killpg(os.getpgid(pid), signal.SIGTERM)  # type: ignore
