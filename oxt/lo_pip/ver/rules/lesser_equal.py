from __future__ import annotations
from typing import List
import re
from ..req_version import ReqVersion
from .ver_rule_base import VerRuleBase


class LesserEqual(VerRuleBase):
    """
    A class to represent a Less than version.
    """

    def _starts_with_less_equal(self, string: str) -> bool:
        """Check if a string starts with <= followed by a space or an integer.

        Args:
            string (str): The input string.

        Returns:
            bool: True if the string matches the pattern, False otherwise.

        Example:

            .. code-block:: python

                >>> self._starts_with_less_equal("<= 1")
                True
                >>> self._starts_with_less_equal("<=  2")
                True
                >>> self._starts_with_less_equal("<=a")
                False
                >>> self._starts_with_less_equal("<1")
                False
        """
        pattern = r"^<=\s*\d"
        return bool(re.match(pattern, string))

    def get_is_match(self) -> bool:
        """Check if the version matches the given string."""
        vlen = len(self.vstr)
        if vlen < 3:
            return False
        if not self._starts_with_less_equal(self.vstr):
            return False
        try:
            versions = self.get_versions()
            if len(versions) == 1:
                return True
            return False
        except Exception:
            return False

    def get_versions(self) -> List[ReqVersion]:
        """Get the list of versions. In this case it will be a single version, unless vstr is invalid in which case it will be an empty list."""
        ver = self.vstr[2:].strip()
        if ver == "":
            return []
        return [ReqVersion(f"<={ver}")]

    def get_versions_str(self) -> str:
        """
        Gets the list of versions as strings.
        In this case in the form of ``<= 1.2.3``.

        Retruns:
            str: The version as a string or an empty string if the version is invalid.
        """
        versions = self.get_versions()
        if len(versions) == 1:
            return versions[0].get_pip_ver_str()
        return ""

    def get_version_is_valid(self, check_version: str) -> int:
        """
        Check if the version is valid. check_version is valid if it is less than or equal to vstr.

        Args:
            check_version (str): Version to check in the form of ``1.2.3`` (no prefix).

        Returns:
            int: ``0`` if the check_version is less than or equal to vstr, ``1`` if the check_version is greater than vstr.
                ``-2`` if the version is invalid.
        """
        try:
            check_ver = ReqVersion(f"=={check_version}")
            versions = self.get_versions()
            if len(versions) != 1:
                return -2
            v1 = versions[0]
            if check_ver < v1:
                return 0
            elif check_ver == v1:
                return 0
            else:
                # greater
                return 1
        except Exception:
            return -2
