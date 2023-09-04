from __future__ import unicode_literals, annotations
from typing import TYPE_CHECKING
import uno
import unohelper
import sys
import os
import logging

from com.sun.star.task import XJob

if TYPE_CHECKING:
    # just for design time
    from lo_pip.config import Config
    from lo_pip.install.install_pip import InstallPip

implementation_name = "___lo_identifier___.___lo_implementation_name___"
implementation_services = ("com.sun.star.task.Job",)

LOG_LEVEL = logging.DEBUG


handler = logging.FileHandler("/home/paul/tmp/py_runner.log", encoding="utf8", delay=True)
handler.setLevel(LOG_LEVEL)

logging.basicConfig(handlers=[handler], level=LOG_LEVEL)

logger = logging.getLogger(__name__)
logger.setLevel(LOG_LEVEL)

# create a file handler

# create a logging format
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
handler.setFormatter(formatter)

# add the handlers to the logger
logger.addHandler(handler)

logger.info("py_runner module imported")


class ___lo_implementation_name___(unohelper.Base, XJob):
    def __init__(self, ctx):
        logger.debug("OooPipRunner Init")
        self.ctx = ctx

    def execute(self, *args):
        # make sure our pythonpath is in sys.path
        logger.debug("OooPipRunner executing")
        this_pth = os.path.dirname(__file__)
        pth_added = False
        try:
            pth = os.path.join(os.path.dirname(__file__), "py_pkgs.zip")
            if not pth in sys.path:
                logger.debug(f"{pth}: appended to sys.path")
                sys.path.append(pth)

            # pythonpath is automatically append b LibreOffice, Can do a check here just in case.
            # pth = os.path.join(os.path.dirname(__file__), "pythonpath")
            # if os.path.exists(pth) and os.path.isdir(pth) and not pth in sys.path:
            #     sys.path.append(pth)

            if not TYPE_CHECKING:
                # run time
                if not this_pth in sys.path:
                    sys.path.append(this_pth)
                    pth_added = True
                    logger.debug(f"{this_pth}: appended to sys.path")
                from lo_pip.config import Config
                from lo_pip.install.install_pip import InstallPip
            logger.debug("Imported Config and InstallPip")
            cfg = Config()
            logger.debug("Created config instance")
            if cfg.py_pkg_dir:
                # add package zip file to the sys.path
                pth = os.path.join(os.path.dirname(__file__), f"{cfg.py_pkg_dir}.zip")

                if os.path.exists(pth) and os.path.isfile(pth) and os.path.getsize(pth) > 0 and not pth in sys.path:
                    logger.debug(f"{pth} appended to sys.path")
                    sys.path.append(pth)
            # sys.path.insert(0, sys.path.pop(sys.path.index(pth)))

            pip_installer = InstallPip()
            logger.debug("Created InstallPip instance")
            if pip_installer.is_pip_installed():
                logger.info("Pip is already installed")
            else:
                logger.info("Pip is not installed. Attempting to install")
                pip_installer.install_pip()
                if pip_installer.is_pip_installed():
                    logger.info("Pip has been installed")
                else:
                    logger.info("Pip was not successfully installed")
            logger.info("OooPipRunner execute Done!")
        except Exception as err:
            logger.error(err)
        finally:
            if pth_added:
                sys.path.remove(this_pth)
        return


g_TypeTable = {}
# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

# add the FormatFactory class to the implementation container,
# which the loader uses to register/instantiate the component.
g_ImplementationHelper.addImplementation(___lo_implementation_name___, implementation_name, implementation_services)
