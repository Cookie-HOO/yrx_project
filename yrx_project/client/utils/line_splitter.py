from PyQt5.QtWidgets import QSplitter, QVBoxLayout, QHBoxLayout, QWidget, QWidgetItem, QSizePolicy
from PyQt5.QtCore import Qt

class LineSplitterWrapper:
    def __init__(self, splitter):
        self.splitter = splitter
        self.design()

    def design(self):
        # 在代码中设置 Splitter 样式
        self.splitter.setStyleSheet("""
            QSplitter::handle {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #D9D9D9, stop:1 #D9D9D9);
                width: 8px;  /* 水平方向手柄宽度 */
                height: 3px; /* 垂直方向手柄高度 */
                margin: 3px;
                image: url(handle_icon.png);  /* 可选：添加图标 */
            }
            QSplitter::handle:hover {
                background: #6B6FA7;
            }
        """)
