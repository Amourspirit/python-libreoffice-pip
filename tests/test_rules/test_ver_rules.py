from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


from oxt.lo_pip.ver.rules.ver_rules import VerRules


@pytest.mark.parametrize(
    "vstr,gstr,valid_str",
    [
        ("==1.2.3", "== 1.2.3", "1.2.3"),
        (">1.2.3", "> 1.2.3", "1.2.4"),
        ("<1.2.3", "< 1.2.3", "1.2.2"),
        (">=1.2.3", ">= 1.2.3", "1.2.4"),
        ("<=1.2.3", "<= 1.2.3", "1.2.3"),
        ("!=1.2.3", "!= 1.2.3", "1.2"),
        ("<> 1.2.3", "!= 1.2.3", "2.2"),
        ("^1.2", ">= 1.2.0 < 2.0.0", "1.5"),
        ("~1.2", ">= 1.2.0 < 1.3.0", "1.2.4"),
        ("2.*", ">= 2.0.0 < 3.0.0", "2.2.4"),
        ("== 1.2.3", "== 1.2.3", "1.2.3"),
        ("> 1.2.3", "> 1.2.3", "1.2.4"),
        ("< 1.2.3", "< 1.2.3", "1.2.2"),
        (">= 1.2.3", ">= 1.2.3", "1.2.4"),
        ("<= 1.2.3", "<= 1.2.3", "1.2.3"),
        ("!= 1.2.3", "!= 1.2.3", "1.2"),
        ("^ 1.2", ">= 1.2.0 < 2.0.0", "1.5"),
        ("~ 1.2", ">= 1.2.0 < 1.3.0", "1.2.4"),
    ],
)
def test_rules(vstr: str, gstr: str, valid_str: str) -> None:
    vr = VerRules()
    rules = vr.get_matched_rules(vstr)
    assert len(rules) == 1
    rule = rules[0]
    assert rule.get_versions_str(vstr) == gstr
    assert rule.get_version_is_valid(valid_str, vstr) == 0
