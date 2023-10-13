from __future__ import annotations
from typing import Any, TYPE_CHECKING

from .run_time_dialog_base import RuntimeDialogBase


class TempDialog(RuntimeDialogBase):
    """
    Temp dialog.

    Be aware this dialog does not have a button.
    """

    MARGIN = 3
    BUTTON_WIDTH = 80
    BUTTON_HEIGHT = 26
    HEIGHT = MARGIN * 3 + BUTTON_HEIGHT * 2
    WIDTH = 400

    def __init__(self, ctx: Any, title: str, msg: str = "Complete", parent: Any = None, align: int = 0):
        """
        Constructor

        Args:
            ctx (Any): Context
            title (str): Dialog title
            msg (str, optional): Dialog Message. Defaults to "Complete".
            parent (Any, optional): Parent Window. Defaults to None.
            align (int, optional): Message Alignment, 0 for left, 1 for center and 2 for right. Defaults to 0.
        """
        super().__init__(ctx)
        self.title = title
        self._msg = msg
        self.parent = parent
        self._is_init = False
        if align < 0 or align > 2:
            align = 0
        self._align = align

    def update(self, msg: str):
        self.dialog.getControl("label").setText(msg)  # type: ignore

    def _init(self):
        if self._is_init:
            return
        margin = self.MARGIN
        self.create_dialog(self.title, size=(self.WIDTH, self.HEIGHT))
        self.create_label(
            name="label",
            command="",
            pos=(margin, margin),
            size=(self.WIDTH - margin * 2, self.HEIGHT - (margin * 2)),
            prop_names=("Label", "MultiLine", "Align"),
            prop_values=(self._msg, True, self._align),
        )
        if self.parent:
            self.dialog.createPeer(self.parent.getToolkit(), self.parent)
        self._is_init = True

    def _result(self):
        return None

    @property
    def msg(self):
        """Gets/Sets message"""
        return self._msg
