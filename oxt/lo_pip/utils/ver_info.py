from __future__ import annotations
from packaging.version import Version
import re


def version_compare(v1: str, v2: str, op: str = "") -> bool:
    """
    Compare two version strings.

    Args:
        v1 (str): First version string.
        v2 (str): Second version string.
        op (str, optional): Compare Operator. Defaults to "".

    Raises:
        ValueError: If the operator is invalid.

    Returns:
        bool: True if the comparison is true, False otherwise.

    Note:
        Operaters are case insensitive and can be any of the following:

        - ``==``, ``=``, ``eq``: Equal
        - ``!=``, ``<>``, ``ne``: Not equal
        - ``>``, ``gt``: Greater than
        - ``>=``, ``ge``: Greater than or equal
        - ``<``, ``lt``: Less than
        - ``<=``, ``le``: Less than or equal
    """
    ver1 = Version(v1)
    ver2 = Version(v2)
    if op == "":
        return ver1 == ver2
    op = op.lower()
    if op in {"=", "==", "eq"}:
        return ver1 == ver2
    if op in {"<", "lt"}:
        return ver1 < ver2
    if op in {"<=", "le"}:
        return ver1 <= ver2
    if op in {">", "gt"}:
        return ver1 > ver2
    if op in {">=", "ge"}:
        return ver1 >= ver2
    if op in {"!=", "<>", "ne"}:
        return ver1 != ver2
    raise ValueError(f"Invalid operator: {op}")


def get_version_prefix(version: str) -> str:
    """Get the version prefix.

    Args:
        version (str): The version string.

    Returns:
        str: The prefix.
    """
    match = re.search(r"\d", version)
    if match:
        prefix = version[: match.start()]
    else:
        prefix = ""
    return prefix


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
            prefix = version[: match.start()]
            ver = version[match.start() :]
        else:
            # no prefix means equal.
            prefix = "=="
            ver = version
        if not self._validate_prefix(prefix):
            raise ValueError(f"Invalid prefix: {prefix}")
        self._prefix = prefix
        return ver

    def _validate_prefix(self, prefix: str) -> bool:
        if prefix == "":
            return True
        return prefix in {"==", "!=", "<>", "<", "<=", ">", ">="}
