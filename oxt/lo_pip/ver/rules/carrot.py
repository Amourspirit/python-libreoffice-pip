from __future__ import annotations
from ..req_version import ReqVersion
from typing import List


class Carrot:
    """
    A class to represent a carrot version.

    Caret requirements allow SemVer compatible updates to a specified version.
    An update is allowed if the new version number does not modify the left-most non-zero digit
    in the major, minor, patch grouping.
    """

    def get_is_match(self, vstr: str) -> bool:
        """Check if the version matches the given string."""
        vlen = len(vstr)
        if vlen < 2 or vlen > 6:
            return False
        return vstr.startswith("^")

    def get_versions(self, vstr: str) -> List[ReqVersion]:
        """Get the list of versions."""
        ver = vstr[1:]
        if ver == "":
            return []
        v1 = ReqVersion(f">={ver}")
        v2 = ReqVersion(f"<{v1.major + 1}.0.0")
        return [v1, v2]

    def get_versions_str(self, vstr: str) -> str:
        """Get the list of versions as strings."""
        versions = self.get_versions(vstr)
        if len(versions) != 2:
            return ""
        v1 = versions[0]
        v2 = versions[1]
        return f"{v1.get_pip_ver_str()} {v2.get_pip_ver_str()}"

    def get_version_is_valid(self, check_version: str, vstr: str) -> int:
        """
        Check if the version is valid.

        Args:
            check_version (str): Version to check in the form of ``1.2.3`` (no prefix).
            vstr (str): version string in the form of ``^1.2.3``.

        Returns:
            int: ``-1`` if the version is less than the range, ``0`` if the version is in the range, ``1`` if the version is greater than the range.
        """
        check_ver = ReqVersion(f"=={check_version}")
        versions = self.get_versions(vstr)
        v1 = versions[0]
        v2 = versions[1]
        if check_ver >= v1 and check_ver < v2:
            return 0
        if check_ver < v1:
            return -1
        return 1
