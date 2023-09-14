from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


from oxt.lo_pip.ver.rules.carrot import Carrot


@pytest.mark.parametrize(
    "match",
    [
        ("^1.0.0"),
        ("^1.0"),
        ("^1"),
        ("^0.0.1"),
    ],
)
def test_is_match(match: str) -> None:
    rule = Carrot()
    assert rule.get_is_match(match)


def test_get_version() -> None:
    rule = Carrot()
    ver = "^1.2.4"
    assert rule.get_is_match(ver)
    versions = rule.get_versions(ver)
    assert len(versions) == 2
    assert versions[0].prefix == ">="
    assert versions[0].major == 1
    assert versions[0].minor == 2
    assert versions[0].micro == 4

    assert versions[1].prefix == "<"
    assert versions[1].major == 2
    assert versions[1].minor == 0
    assert versions[1].micro == 0

    pip_ver_str = rule.get_versions_str(ver)
    assert pip_ver_str == ">=1.2.4, <2.0.0"

    ver = "^1.2.dev1"
    assert rule.get_is_match(ver)
    versions = rule.get_versions(ver)
    assert len(versions) == 2
    assert versions[0].prefix == ">="
    assert versions[0].major == 1
    assert versions[0].minor == 2
    assert versions[0].dev == 1
    assert versions[0].is_devrelease

    assert versions[1].prefix == "<"
    assert versions[1].major == 2
    assert versions[1].minor == 0
    assert versions[1].micro == 0

    pip_ver_str = rule.get_versions_str(ver)
    assert pip_ver_str == ">=1.2.dev1, <2.0.0"

    ver = "^ 1.2.pre1"
    assert rule.get_is_match(ver)
    versions = rule.get_versions(ver)
    assert len(versions) == 2
    assert versions[0].prefix == ">="
    assert versions[0].major == 1
    assert versions[0].minor == 2
    assert versions[0].is_prerelease

    assert versions[1].prefix == "<"
    assert versions[1].major == 2
    assert versions[1].minor == 0
    assert versions[1].micro == 0

    pip_ver_str = rule.get_versions_str(ver)
    assert pip_ver_str == ">=1.2rc1, <2.0.0"

    ver = "^1.2.rc1"
    assert rule.get_is_match(ver)
    versions = rule.get_versions(ver)
    assert len(versions) == 2
    assert versions[0].prefix == ">="
    assert versions[0].major == 1
    assert versions[0].minor == 2
    assert versions[0].is_prerelease

    assert versions[1].prefix == "<"
    assert versions[1].major == 2
    assert versions[1].minor == 0
    assert versions[1].micro == 0

    pip_ver_str = rule.get_versions_str(ver)
    assert pip_ver_str == ">=1.2rc1, <2.0.0"

    ver = "^1.2rc1"
    assert rule.get_is_match(ver)
    versions = rule.get_versions(ver)
    assert len(versions) == 2
    assert versions[0].prefix == ">="
    assert versions[0].major == 1
    assert versions[0].minor == 2
    assert versions[0].is_prerelease

    assert versions[1].prefix == "<"
    assert versions[1].major == 2
    assert versions[1].minor == 0
    assert versions[1].micro == 0

    pip_ver_str = rule.get_versions_str(ver)
    assert pip_ver_str == ">=1.2rc1, <2.0.0"


def test_get_version_is_valid() -> None:
    rule = Carrot()
    ver = "^1.2.4"
    assert rule.get_is_match(ver)
    assert rule.get_version_is_valid("1.2.4", ver) == 0
    assert rule.get_version_is_valid("1.2.5", ver) == 0
    assert rule.get_version_is_valid("1.3.0", ver) == 0
    assert rule.get_version_is_valid("2.0.0", ver) == 1
    assert rule.get_version_is_valid("1.2.3", ver) == -1


@pytest.mark.parametrize(
    "check_ver,vstr,result",
    [
        ("1.2.4", "^1.2.4", 0),
        ("1.2.5", "^ 1.2.4", 0),
        ("1.2rc1", "^1.2.pre1", 0),
        ("1.2rc1", "^1.2.rc1", 0),
        ("1.2rc1", "^1.2.pre2", -1),
        ("1.2rc2", "^1.2.pre2", 0),
        ("1.2dev3", "^1.2.dev3", 0),
        ("1.2dev1", "^1.2dev2", -1),
        ("1.2dev3", "^1.2dev2", 0),
        ("1.2post3", "^1.2.post3", 0),
        ("1.2post1", "^1.2post2", -1),
        ("2.2post3", "^1.2post2", 1),
    ],
)
def test_get_version_is_valid_suffix(check_ver: str, vstr: str, result: int) -> None:
    rule = Carrot()
    assert rule.get_is_match(vstr)
    assert rule.get_version_is_valid(check_ver, vstr) == result
