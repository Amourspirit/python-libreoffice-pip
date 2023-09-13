from __future__ import annotations
from ..req_version import ReqVersion
from typing import Protocol, List


class VerProto(Protocol):
    """A protocol for version objects."""

    # def __call__(self) -> "VerProto":
    #     ...

    def get_is_match(self, vstr: str) -> bool:
        """Check if the version matches the given string."""
        ...

    def get_versions(self, vstr: str) -> List[ReqVersion]:
        """Get the list of versions."""
        ...

    def get_versions_str(self, vstr: str) -> str:
        """Get the list of versions as strings."""
        ...

    def get_version_is_valid(self, check_version: str, vstr: str) -> int:
        """
        Check if the version is valid.

        Args:
            check_version (str): Version to check in the form of ``1.2.3`` (no prefix).
            vstr (str): version string in the form of ``^1.2.3`` or other prefix.

        Returns:
            Returns:
            int: ``0`` if the version is in the range, some other value if the version is not in the range.
                Implemented by each rule.
        """
        ...
