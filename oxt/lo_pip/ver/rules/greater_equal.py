from __future__ import annotations
from ..req_version import ReqVersion
from typing import List
import re


class GreaterEqual:
    """
    A class to represent a Greater than or equal to version.
    """

    def _starts_with_greater_equal(self, string: str) -> bool:
        """Check if a string starts with >= followed by a space or an integer.

        Args:
            string (str): The input string.

        Returns:
            bool: True if the string matches the pattern, False otherwise.

        Example:

            .. code-block:: python

                >>> self._starts_with_greater_equal(">= 1")
                True
                >>> self._starts_with_greater_equal(">=  2")
                True
                >>> self._starts_with_greater_equal(">=a")
                False
                >>> self._starts_with_greater_equal(">1")
                False
        """
        pattern = r"^>=\s*\d"
        return bool(re.match(pattern, string))

    def get_is_match(self, vstr: str) -> bool:
        """Check if the version matches the given string."""
        vlen = len(vstr)
        if vlen < 3 or vlen > 10:
            return False
        return self._starts_with_greater_equal(vstr)

    def get_versions(self, vstr: str) -> List[ReqVersion]:
        """Get the list of versions. In this case it will be a single version, unless vstr is invalid in which case it will be an empty list."""
        ver = vstr[2:].strip()
        if ver == "":
            return []
        return [ReqVersion(f">={ver}")]

    def get_versions_str(self, vstr: str) -> str:
        """
        Gets the list of versions as strings.
        In this case in the form of ``>= 1.2.3``.

        Retruns:
            str: The version as a string or an empty string if the version is invalid.
        """
        versions = self.get_versions(vstr)
        if len(versions) == 1:
            return versions[0].get_pip_ver_str()
        return ""

    def get_version_is_valid(self, check_version: str, vstr: str) -> int:
        """
        Check if the version is valid. check_version is valid if it is greater than or equal the vstr.

        Args:
            check_version (str): Version to check in the form of ``1.2.3`` (no prefix).
            vstr (str): version string in the form of ``>=1.2.2``.

        Returns:
            int: ``0`` if the check_version is greater than or equal to the vstr, ``-1`` if the check_version is less than vstr.
                ``-2`` if the version is invalid.
        """
        try:
            check_ver = ReqVersion(f"=={check_version}")
            versions = self.get_versions(vstr)
            if len(versions) != 1:
                return -2
            v1 = versions[0]
            if check_ver > v1:
                return 0
            elif check_ver == v1:
                return 0
            else:
                # less than
                return -1
        except Exception:
            return -2
