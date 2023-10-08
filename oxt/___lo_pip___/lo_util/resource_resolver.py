from __future__ import annotations
import uno
from typing import Any, cast, TYPE_CHECKING

from ..config import Config
from ..oxt_logger import OxtLogger

from com.sun.star.lang import Locale, IllegalArgumentException

if TYPE_CHECKING:
    from com.sun.star.util import PathSubstitution  # service
    from com.sun.star.resource import StringResourceWithLocation  # service


class ResourceResolver:
    """Resource Resolver for localized strings"""

    def __init__(self, ctx: Any):
        self._logger = OxtLogger(log_name=__name__)
        try:
            self.ctx = ctx
            self.service_manager = self.ctx.getServiceManager()
            self.locale = self._get_env_locale()
            self.resource_resolver = self._get_resource_resolver()
            self.version = self._get_ext_ver()
        except Exception as err:
            self._logger.error(f"ResourceResolver.__init__: {err}", exc_info=True)

    def _get_ext_ver(self) -> str:
        """Get addon version number"""
        config = Config()
        info = config.extension_info.get_extension_info(config.lo_implementation_name)
        try:
            return info[1]
        except Exception:
            return "0.0.0"

    def _get_env_locale(self):
        """Get interface locale"""
        ps = cast(
            "PathSubstitution",
            self.service_manager.createInstanceWithContext("com.sun.star.util.PathSubstitution", self.ctx),
        )
        v_lang = ps.getSubstituteVariableValue("vlang")
        a_lang = v_lang.split("-") + 2 * [""]
        return Locale(*a_lang[:3])

    def _get_resource_resolver(self) -> StringResourceWithLocation:
        # url = self._get_ext_path() + "python"
        config = Config()
        url = f"vnd.sun.star.extension://{config.lo_identifier}/{config.resource_dir_name}"
        handler = self.service_manager.createInstanceWithContext("com.sun.star.task.InteractionHandler", self.ctx)
        return cast(
            "StringResourceWithLocation",
            self.service_manager.createInstanceWithArgumentsAndContext(
                "com.sun.star.resource.StringResourceWithLocation",
                (url, False, self.locale, config.resource_properties_prefix, "", handler),
                self.ctx,
            ),
        )

    def resolve_string(self, id: str) -> str:
        """Resolve localized string

        Args:
            id (str): resource id

        Returns:
            str: localized string
        """
        try:
            return self.resource_resolver.resolveString(id)
        except Exception as err:
            self._logger.error(f"ResourceResolver.resolve_string: {err}", exc_info=True)
            raise
