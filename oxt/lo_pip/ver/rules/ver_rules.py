from __future__ import annotations
from typing import List, Type, Dict
from .ver_proto import VerProto
from .carrot import Carrot
from .equals import Equals
from .greater import Greater
from .greater_equal import GreaterEqual
from .greater import Greater
from .lesser_equal import LesserEqual
from .lesser import Lesser
from .not_equals import NotEquals
from .tilde import Tilde
from .wildcard import Wildcard

# https://www.darius.page/pipdev/


class VerRules:
    """Manages rules for Versions"""

    def __init__(self) -> None:
        self._rules: List[Type[VerProto]] = []
        self._cache: Dict[str, VerProto] = {}
        self._register_known_rules()

    def __len__(self) -> int:
        return len(self._rules)

    # region Methods

    def register_rule(self, rule: Type[VerProto]) -> None:
        """
        Register rule

        Args:
            rule (VerProto): Rule to register
        """
        if rule in self._rules:
            # msg = f"{self.__class__.__name__}.register_rule() Rule is already registered"
            # log.logger.warning(msg)
            return
        self._reg_rule(rule=rule)

    def unregister_rule(self, rule: Type[VerProto]):
        """
        Unregisters Rule

        Args:
            rule (VerProto): Rule to unregister

        Raises:
            ValueError: If an error occurs
        """
        try:
            key = str(id(rule))
            if key in self._cache:
                del self._cache[key]
            self._rules.remove(rule)
        except ValueError as e:
            msg = f"{self.__class__.__name__}.unregister_rule() Unable to unregister rule."
            raise ValueError(msg) from e

    def _reg_rule(self, rule: Type[VerProto]):
        self._rules.append(rule)

    def _register_known_rules(self):
        self._reg_rule(rule=Carrot)
        self._reg_rule(rule=Equals)
        self._reg_rule(rule=Greater)
        self._reg_rule(rule=GreaterEqual)
        self._reg_rule(rule=Lesser)
        self._reg_rule(rule=LesserEqual)
        self._reg_rule(rule=NotEquals)
        self._reg_rule(rule=Tilde)
        self._reg_rule(rule=Wildcard)

    def get_matched_rules(self, vstr: str) -> List[VerProto]:
        """
        Get matched rules

        Args:
            vstr (str): Version in string form, e.g. ``==1.2.3``

        Returns:
            List[VerProto]: List of matched rules
        """
        resuts: List[VerProto] = []
        for rule in self._rules:
            key = str(id(rule))
            if key in self._cache:
                inst = self._cache[key]
            else:
                inst = rule()
                self._cache[key] = inst
            if inst.get_is_match(vstr):
                resuts.append(inst)
        return resuts

    # endregion Methods
