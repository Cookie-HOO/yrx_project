import typing
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QLabel, QGridLayout
from PyQt5.QtCore import Qt


class GridLayout:
    def __init__(self, grid_layout: QGridLayout, max_cols: int):
        self.grid_layout = grid_layout
        self.max_cols = max_cols
        self._cur_row = 0
        self._cur_col = 0

    def add_pic(self, pic_or_path: typing.Union[str, QPixmap]):
        """添加图片到网格布局，自动换行"""
        # 加载图片
        pixmap = pic_or_path if isinstance(pic_or_path, QPixmap) else QPixmap(pic_or_path)
        if pixmap.isNull():
            return

        # 动态计算尺寸
        container = self.grid_layout.parentWidget()
        if container and container.width() > 0:
            # 计算可用宽度（扣除边距和间距）
            margins = self.grid_layout.getContentsMargins()
            spacing = self.grid_layout.spacing()
            available_width = container.width() - (margins[0] + margins[2]) - (spacing * (self.max_cols - 1))
            cell_width = max(100, available_width // self.max_cols)
            height = int(cell_width * pixmap.height() / pixmap.width())
        else:
            cell_width, height = 200, 150  # 默认尺寸

        # 创建图片标签
        label = QLabel()
        label.setPixmap(pixmap.scaled(
            cell_width,
            height,
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        ))
        label.setAlignment(Qt.AlignCenter)

        # 添加到布局并更新位置
        self.grid_layout.addWidget(label, self._cur_row, self._cur_col)
        self._cur_col += 1
        if self._cur_col >= self.max_cols:
            self._cur_col = 0
            self._cur_row += 1

    def set_max_cols(self, max_cols: int):
        """动态调整最大列数并重新排列布局"""
        # 保存现有控件
        widgets = [
            self.grid_layout.itemAt(i).widget()
            for i in range(self.grid_layout.count())
            if self.grid_layout.itemAt(i).widget()
        ]

        # 清空并重置布局
        self.clear()
        self.max_cols = max_cols

        # 重新排列控件
        row = col = 0
        for widget in widgets:
            self.grid_layout.addWidget(widget, row, col)
            col += 1
            if col >= self.max_cols:
                col = 0
                row += 1

        # 更新当前插入位置
        total = len(widgets)
        self._cur_row = total // self.max_cols
        self._cur_col = total % self.max_cols

    def clear(self):
        """清空布局并重置状态"""
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            if widget := item.widget():
                widget.deleteLater()
        self._cur_row = 0
        self._cur_col = 0
