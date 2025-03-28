from PyQt5.QtWidgets import QSplitter, QVBoxLayout, QWidgetItem
from PyQt5.QtCore import Qt


class LineSplitterWrapper:
    def __init__(self, line_widget):
        self.line_widget = line_widget
        self.replace_line_with_splitter()

    def replace_line_with_splitter(self):
        # 1. 获取self.line的父布局
        line = self.line_widget
        parent_layout = line.parent().layout()  # 获取线所在布局[[5]][[7]]

        # 2. 找到self.line在布局中的索引
        line_index = parent_layout.indexOf(line)
        if line_index == -1:
            raise ValueError("Line not found in layout")

        # 3. 获取前后部件（索引-1为前面，索引+1为后面）
        upper_index = line_index - 1
        lower_index = line_index + 1

        # 4. 检查前后部件是否存在
        if upper_index < 0 or lower_index >= parent_layout.count():
            raise IndexError("Invalid layout structure")

        # 5. 获取具体部件（注意：QLayoutItem需要转换为QWidget）
        upper_item = parent_layout.itemAt(upper_index)
        lower_item = parent_layout.itemAt(lower_index)

        upper_widget = upper_item.widget()
        lower_widget = lower_item.widget()

        # 6. 移除原有部件
        parent_layout.removeItem(upper_item)  # 移除上部件[[5]]
        parent_layout.removeItem(lower_item)  # 移除下部件
        parent_layout.removeWidget(line)  # 移除线

        # 7. 创建垂直方向的QSplitter
        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(upper_widget)
        splitter.addWidget(lower_widget)

        # 8. 将QSplitter添加到原线的位置
        parent_layout.insertWidget(line_index, splitter)  # 替换线的位置[[8]]

        # 9. 清理旧部件
        line.setParent(None)  # 彻底移除旧线[[5]]