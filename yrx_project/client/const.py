from PyQt5.QtGui import QColor

from yrx_project.const import *

WINDOW_INIT_SIZE = (1068, 681)
WINDOW_INIT_DISTANCE = (800, 200)
WINDOW_TITLE = "catfisher"
STATIC_FILE_PATH = os.path.join(STATIC_PATH, "{file}")
UI_PATH = os.path.join(ROOT_IN_EXE_PATH, 'ui', "{file}") if is_prod else os.path.join(os.path.dirname(__file__), "ui", "{file}")

# color
COLOR_WHITE = QColor(255, 255, 255)
COLOR_RED = QColor(245, 184, 184)
COLOR_GREEN = QColor(199, 242, 174)
COLOR_YELLOW = QColor(250, 243, 160)
COLOR_BLUE = QColor(206, 220, 245)

# color emoji
COLOR_STR_RED = "🟥"
COLOR_STR_YELLOW = "🟨"
COLOR_STR_GREEN = "🟩"
COLOR_STR_BLUE = "🟦"

# item type
EDITABLE_TEXT = "editable_text"
EDITABLE_INT = "editable_int"
EDITABLE_COLOR = "editable_color"
READONLY_TEXT = "readonly_text"
DROPDOWN = "dropdown"

READONLY_VALUE = "---"

