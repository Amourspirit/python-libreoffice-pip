import logging
import sys
from logging import Logger
from logging.handlers import TimedRotatingFileHandler

from ..config import Config


class OxtLogger(Logger):
    def __init__(self, log_file: str = "", log_name: str = "", *args, **kwargs):
        self.config = Config()
        self.formatter = logging.Formatter(self.config.log_format)
        if not log_file:
            log_file = self.config.log_file
        self.log_file = log_file
        if not log_name:
            log_name = self.config.log_name
        self.log_name = log_name

        # Logger.__init__(self, name=log_name, level=cfg.log_level)
        super().__init__(name=log_name, level=self.config.log_level)

        if self.log_file:
            self.addHandler(self.get_file_handler())
        else:
            self.addHandler(self.get_console_handler())

        # with this pattern, it's rarely necessary to propagate the| error up to parent
        self.propagate = False

    def get_console_handler(self):
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(self.formatter)
        return console_handler

    def get_file_handler(self):
        log_file = self.log_file
        file_handler = TimedRotatingFileHandler(
            log_file, when="W0", interval=1, backupCount=3, encoding="utf8", delay=True
        )
        # file_handler = logging.FileHandler(log_file, mode="w", encoding="utf8", delay=True)
        file_handler.setFormatter(self.formatter)
        file_handler.setLevel(self.config.log_level)
        return file_handler
