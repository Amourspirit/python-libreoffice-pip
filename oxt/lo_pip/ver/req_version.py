from __future__ import annotations
from packaging.version import Version
import re


class ReqVersion(Version):
    """A class to represent a version requirement."""

    def __init__(self, version: str) -> None:
        """Initialize a Version object.

        Args:
            version (str): The string representation of a version which will be parsed and normalized
                before use.

        Raises:
            InvalidVersion: If the ``version`` does not conform to PEP 440 in any way then this exception will be raised.
        """
        ver = self._process_full_verion(version)
        super().__init__(ver)

    def _process_full_verion(self, version: str) -> str:
        match = re.search(r"\d", version)
        if match:
            prefix = version[: match.start()].strip()
            ver = version[match.start() :].strip()
        else:
            # no prefix means equal.
            prefix = "=="
            ver = version.strip()
        if not self._validate_prefix(prefix):
            raise ValueError(f"Invalid prefix: {prefix}")
        self._prefix = prefix
        return ver

    def _validate_prefix(self, prefix: str) -> bool:
        if prefix == "":
            return True
        return prefix in {"==", "!=", "<>", "<", "<=", ">", ">="}

    def get_ver_is_valid(self, version: str) -> bool:
        """Check if the version is valid."""
        ver = Version(version)
        if self._prefix == "==":
            return ver == self
        if self._prefix == "!=":
            return ver != self
        if self._prefix == "<":
            return ver < self
        if self._prefix == "<=":
            return ver <= self
        if self._prefix == ">":
            return ver > self
        if self._prefix == ">=":
            return ver >= self
        return False

    def get_pip_ver_str(self) -> str:
        """Get the pip version string."""
        return self._prefix + str(self)
        # return f"{self._prefix} {self.major}.{self.minor}.{self.micro}"

    @property
    def prefix(self) -> str:
        """Get the prefix such as ``==`` or ``>=``."""
        return self._prefix
