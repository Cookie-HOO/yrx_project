from PyQt5.QtWidgets import QApplication, QToolButton, QColorDialog, QMenu
from PyQt5.QtGui import QColor, QPixmap, QIcon, QPalette
from PyQt5.QtCore import Qt, pyqtSignal, QSize


class ColorPickerToolButton(QToolButton):
    colorChanged = pyqtSignal(QColor)

    def __init__(self, initial_color=QColor(Qt.white), parent=None):
        super().__init__(parent)
        self.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        # self.setMinimumSize(1, 28)
        self.set_color(initial_color)
        self.clicked.connect(self.choose_color)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)

        # 固定样式初始化（不随颜色改变）
        base_style = """
            QToolButton {{
                border: null;
                border-radius: 3px;
                padding: 4px 8px;
                text-align: left;
                background: palette(base);  /* 保持默认背景色 */
                color: {text_color};  /* 初始文本颜色 */
            }}
            QToolButton::menu-indicator {{ image: none; }}
        """.format(text_color="black")  # 初始文本颜色
        self.setStyleSheet(base_style)

    def set_color(self, color):
        self._color = color

        # 创建颜色块图标
        pixmap = QPixmap(20, 20)
        pixmap.fill(color)
        self.setIcon(QIcon(pixmap))
        self.setIconSize(QSize(20, 20))

        # 设置文本
        self.setText(color.name())

        # 根据颜色亮度动态调整文本颜色（仅修改文本颜色）
        # text_color = "white" if color.lightness() < 128 else "black"
        # self.setStyleSheet(self.styleSheet() + f"""
        #     QToolButton {{
        #         color: {text_color};
        #     }}
        # """)
        self.colorChanged.emit(color)

    def choose_color(self):
        color = QColorDialog.getColor(self._color, self, "选择颜色")
        if color.isValid():
            self.set_color(color)

    @property
    def color(self):
        return self._color

    def show_context_menu(self, pos):
        menu = QMenu()
        copy_action = menu.addAction("复制颜色值")
        paste_action = menu.addAction("粘贴颜色值")
        action = menu.exec_(self.mapToGlobal(pos))

        if action == copy_action:
            QApplication.clipboard().setText(self.color.name())
        elif action == paste_action:
            color = QColor(QApplication.clipboard().text())
            if color.isValid():
                self.set_color(color)