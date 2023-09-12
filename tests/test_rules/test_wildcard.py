from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


from oxt.lo_pip.ver.rules.wildcard import Wildcard


@pytest.mark.parametrize(
    "match",
    [
        ("1.0.*"),
        ("1.*"),
        ("*"),
    ],
)
def test_is_match(match: str) -> None:
    rule = Wildcard()
    assert rule.get_is_match(match)


def test_get_version() -> None:
    rule = Wildcard()
    ver = "1.2.*"
    assert rule.get_is_match(ver)
    versions = rule.get_versions(ver)
    assert len(versions) == 2
    assert versions[0].prefix == ">="
    assert versions[0].major == 1
    assert versions[0].minor == 2
    assert versions[0].micro == 0

    assert versions[1].prefix == "<"
    assert versions[1].major == 1
    assert versions[1].minor == 3
    assert versions[1].micro == 0

    pip_ver_str = rule.get_versions_str(ver)
    assert pip_ver_str == ">= 1.2.0 < 1.3.0"

    ver = "2.*"
    assert rule.get_is_match(ver)
    versions = rule.get_versions(ver)
    assert len(versions) == 2
    assert versions[0].prefix == ">="
    assert versions[0].major == 2
    assert versions[0].minor == 0
    assert versions[0].micro == 0

    assert versions[1].prefix == "<"
    assert versions[1].major == 3
    assert versions[1].minor == 0
    assert versions[1].micro == 0

    pip_ver_str = rule.get_versions_str(ver)
    assert pip_ver_str == ">= 2.0.0 < 3.0.0"

    ver = "*"
    assert rule.get_is_match(ver)
    versions = rule.get_versions(ver)
    assert len(versions) == 1
    assert versions[0].prefix == ">="
    assert versions[0].major == 0
    assert versions[0].minor == 0
    assert versions[0].micro == 0

    pip_ver_str = rule.get_versions_str(ver)
    assert pip_ver_str == ">= 0.0.0"


def test_get_version_is_valid() -> None:
    rule = Wildcard()
    ver = "1.2.*"
    assert rule.get_is_match(ver)
    assert rule.get_version_is_valid("1.2.4", ver) == 0
    assert rule.get_version_is_valid("1.2.9", ver) == 0
    assert rule.get_version_is_valid("1.2.0", ver) == 0
    assert rule.get_version_is_valid("1.1.0", ver) == -1
    assert rule.get_version_is_valid("1.3", ver) == 1

    ver = "1.*"
    assert rule.get_is_match(ver)
    assert rule.get_version_is_valid("1.2.4", ver) == 0
    assert rule.get_version_is_valid("1.2.9", ver) == 0
    assert rule.get_version_is_valid("1.9.9", ver) == 0
    assert rule.get_version_is_valid("0.2.9", ver) == -1
    assert rule.get_version_is_valid("2.0.0", ver) == 1

    ver = "*"
    assert rule.get_is_match(ver)
    assert rule.get_version_is_valid("1.2.4", ver) == 0
    assert rule.get_version_is_valid("0.0.1", ver) == 0
    assert rule.get_version_is_valid("99.0.1", ver) == 0
