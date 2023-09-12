from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


from oxt.lo_pip.utils.ver_info import version_compare


@pytest.mark.parametrize(
    "v1,v2,op,result",
    [
        ("1.0.0", "1.0.0", "", True),
        ("1.0.0", "1.0.0", "==", True),
        ("1.0.0", "1.0.0", "eq", True),
        ("1.0.0", "1.0.0", "EQ", True),
        ("1.0.0", "1.0.0", "=", True),
        ("1.0.0", "1.0.1", "==", False),
        ("1.0.0", "1.0.1", ">=", False),
        ("1.0.0", "1.0.1", "<=", True),
        ("0.1.0", "1.0.1", "<", True),
        ("0.1.0", "1.0.1", "gt", False),
        ("0.1.0", "1.0.1", "ne", True),
        ("0.1.0", "1.0.1", "<>", True),
    ],
)
def test_version(v1, v2, op, result: bool) -> None:
    assert version_compare(v1, v2, op) == result
