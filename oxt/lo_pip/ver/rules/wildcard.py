from __future__ import annotations
from typing import List
import re
from ..req_version import ReqVersion


class Wildcard:
    """
    A class to represent a Wildcard version.

    Wildcard requirements allow for the latest (dependency dependent) version where the wildcard is positioned.
    """

    def _starts_with_equal(self, string: str) -> bool:
        """Check if a string starts with == followed by a space or an integer.

        Args:
            string (str): The input string.

        Returns:
            bool: True if the string matches the pattern, False otherwise.

        Example:

            .. code-block:: python

                >>> self._starts_with_equal("== 1")
                True
                >>> self._starts_with_equal("==  2")
                True
                >>> self._starts_with_equal("<a")
                False
                >>> self._starts_with_equal("<=1")
                False
        """
        pattern = r"^==\s*"
        return bool(re.match(pattern, string))

    def get_is_match(self, vstr: str) -> bool:
        """Check if the version matches the given string."""
        vlen = len(vstr)
        if vlen == 0:
            return False
        if not self._starts_with_equal(vstr):
            return False
        if not vstr.endswith("*"):
            return False
        try:
            versions = self.get_versions(vstr)
            if len(versions) >= 1:
                return True
            return False
        except Exception:
            return False

    def get_versions(self, vstr: str) -> List[ReqVersion]:
        """Get the list of versions."""

        ver = vstr[2:].strip()  # remove ==

        if ver[:-1].strip() == "":
            return [ReqVersion(f">=0.0.0")]
        ver = ver[:-2]  # remove .*
        v1 = ReqVersion(f">={ver}")
        if v1.minor > 0:
            v2 = ReqVersion(f"<{v1.major}.{v1.minor + 1}.0")
        else:
            v2 = ReqVersion(f"<{v1.major + 1}.0.0")
        return [v1, v2]

    def get_versions_str(self, vstr: str) -> str:
        """Get the list of versions as strings."""
        versions = self.get_versions(vstr)
        if len(versions) == 1:
            return versions[0].get_pip_ver_str()
        if len(versions) != 2:
            return ""
        v1 = versions[0]
        v2 = versions[1]
        return f"{v1.get_pip_ver_str()}, {v2.get_pip_ver_str()}"

    def get_version_is_valid(self, check_version: str, vstr: str) -> int:
        """
        Check if the version is valid.

        Args:
            check_version (str): Version to check in the form of ``1.2.3`` (no prefix).
            vstr (str): version string in the form of ``1.2.*``.

        Returns:
            int: ``-1`` if the version is less than the range, ``0`` if the version is in the range, ``1`` if the version is greater than the range.
        """
        check_ver = ReqVersion(f"=={check_version}")
        versions = self.get_versions(vstr)
        if len(versions) == 1:
            # in this instance a sigle version is returned, with a value of >=0.0.0
            return 0
        v1 = versions[0]
        v2 = versions[1]
        if check_ver >= v1 and check_ver < v2:
            return 0
        if check_ver < v1:
            return -1
        return 1
