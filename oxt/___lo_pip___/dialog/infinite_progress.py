from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import threading
import time
import uno

from .run_time_dialog_base import RuntimeDialogBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlEdit  # service


class InfiniteProgressDialog(RuntimeDialogBase):
    """Progress dialog."""

    MARGIN = 3
    BUTTON_WIDTH = 80
    BUTTON_HEIGHT = 26
    HEIGHT = MARGIN * 3 + BUTTON_HEIGHT * 2
    WIDTH = 300

    def __init__(self, ctx: Any, title: str, msg: str = "Please wait...", parent: Any = None):
        super().__init__(ctx)
        self.title = title
        self._msg = msg
        self.parent = parent
        self._is_init = False

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
            size=(self.WIDTH - margin * 2, self.BUTTON_HEIGHT),
            prop_names=("Label",),
            prop_values=(self._msg,),
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

    def set_msg(self, msg: str):
        self._msg = msg
        if self._is_init:
            self.update(msg)


class InfiniteProgress(threading.Thread):
    """
    Infinite progress thread.

    Example:

        ..code-block:: python

            my_progress = InfiniteProgress(self.ctx)
            my_progress.start()
            time.sleep(15)
            my_progress.stop()
    """

    def __init__(self, ctx: Any, title: str = "Infinite Progress", msg: str = "Please wait"):
        super().__init__()
        self._ctx = ctx
        self._title = title
        self._msg = msg
        self._ellipsis = 0
        self._stop_event = threading.Event()

    def run(self):
        # use the setVisible method and not Execute.
        # This makes visible the dialog and makes it not modal.
        in_progress = InfiniteProgressDialog(self._ctx, self._title)
        # in_progress.dialog.setVisible(True)
        # _ = in_progress.execute()
        while not self._stop_event.is_set():
            # code for the process you want to run on the thread
            # ...
            # possibly sleep and update a label to show progress
            self._ellipsis += 1
            in_progress.dialog.setVisible(True)
            in_progress.update(f"{self._msg}{'.' * self._ellipsis}")
            if self._ellipsis > 10:
                self._ellipsis = 1
            time.sleep(1)
        in_progress.dialog.dispose()

    def stop(self):
        self._stop_event.set()
