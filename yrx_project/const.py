# 路径
import os
import sys

is_prod = os.environ.get("env") == 'prod'

# path
PROJECT_PATH = os.getcwd() if is_prod else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ALL_DATA_PATH = os.path.join(PROJECT_PATH, "data")
LOGGER_FILE_PATH = os.path.join(PROJECT_PATH, "logger.log")
ROOT_IN_EXE_PATH = sys._MEIPASS if is_prod else PROJECT_PATH
STATIC_PATH = os.path.join(ROOT_IN_EXE_PATH, "static")
TEMP_PATH = os.path.join(PROJECT_PATH, ".catfisher_temp")  # 临时路径的处理：使用tempfile，程序结束时清理

# formatter
DATE_FORMATTER = '%Y-%m-%d'
TIME_FORMATTER = '%Y-%m-%d %H.%M.%S'
FULL_TIME_FORMATTER = '%Y-%m-%d %H.%M.%S.%f'
DATE_NUM_FORMATTER = '%Y%m%d'

# num
POSITIVE_NUM_CHAR_MAPPING = {1: "一", 2: "二", 3: "三", 4: "四"}
