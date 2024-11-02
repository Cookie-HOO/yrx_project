import datetime
import traceback

from yrx_project.const import LOGGER_FILE_PATH, FULL_TIME_FORMATTER


class Logger:
    def __init__(self, logger_path):
        self.path = logger_path
        self.formatter = FULL_TIME_FORMATTER

    def __write(self, level, msg):
        now = datetime.datetime.now()
        with open(self.path, "a") as f:
            f.write(f"{now.strftime(self.formatter)} [{level}]: {msg}")

    def info(self, msg):
        self.__write("INFO", msg=msg)

    def warn(self, msg):
        self.__write("WARN", msg=msg)

    def error(self, msg):
        self.__write("ERROR", msg=msg)

    def sys_error(self):
        error_info = traceback.format_exc()
        if error_info:
            self.error(error_info)


logger = Logger(LOGGER_FILE_PATH)


def logger_sys_error(func):
    def inner(*args, **kwargs):
        try:
            func(*args, **kwargs)
        except Exception as e:
            logger.sys_error()
            raise e
    return inner

