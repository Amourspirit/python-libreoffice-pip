from __future__ import annotations
from typing import Any, Dict, Tuple
import uno
from pathlib import Path

from ..meta.singleton import Singleton


class Util(metaclass=Singleton):
    def __init__(self) -> None:
        self._ctx = uno.getComponentContext()
        self._service_manager = self._ctx.getServiceManager()  # type: ignore

    def create_instance(self, name: str, with_context: bool = False, args: Any = None) -> Any:
        if with_context:
            return self._service_manager.createInstanceWithContext(name, self._ctx)
        elif args:
            return self._service_manager.createInstanceWithArguments(name, (args,))
        else:
            return self._service_manager.createInstance(name)

    def config(self, name="Work"):
        """
        Return the path name in config
        http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XPathSettings.html

        Examples:

            ``config("module")``
            ``/usr/lib/libreoffice/program``

            ``config("Work")``
            ``/home/user/Documents``
        """
        path = self.create_instance("com.sun.star.util.PathSettings")
        return self.to_system(getattr(path, name))

    def to_system(self, path: str) -> str:
        if path.startswith("file://"):
            path = str(Path(uno.fileUrlToSystemPath(path)).resolve())
        return path
