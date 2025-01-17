from __future__ import annotations
from ___lo_pip___.debug.break_mgr import BreakMgr

break_mgr = BreakMgr()
break_mgr.add_breakpoint("oxt.pythonpath.pip_hello_world")


def hello_world() -> None:
    break_mgr.check_breakpoint("oxt.pythonpath.pip_hello_world")
    hello = "Hello"
    world = "World!"
    print(f"{hello} {world}")
