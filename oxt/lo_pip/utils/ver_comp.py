from __future__ import annotations
from packaging.version import Version


def version_compare(v1: str, v2: str, op: str = "") -> bool:
    """Compare versions."""
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
