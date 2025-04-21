import os

from yrx_project.const import TEMP_PATH

SCENE_NAME = "split_table"
SCENE_TEMP_PATH = str(os.path.join(TEMP_PATH, SCENE_NAME))

REORDER_COLS = ["序号"]  # 需要reorder的列名
