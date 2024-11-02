from PyQt5.QtGui import QPen
from PyQt5.QtWidgets import QGraphicsScene


class GraphWidgetWrapper:
    """画图组件的包装，用于复用画图的方法"""
    def __init__(self, widget):
        self.graph_widget = widget
        self.scene = QGraphicsScene()
        self.x_all = self.graph_widget.width()
        self.y_all = self.graph_widget.height()
        self.y_middle = self.y_all / 2

        self.scene.setSceneRect(0, 0, self.graph_widget.width(), self.graph_widget.height())
        self.y_tick_range = None
        self.height_between = None  # 高度间隔

    def set_y_tick(self, q_pen: QPen, limit=True):
        """设置y轴的tick刻度"""
        if not limit or not self.y_tick_range:
            return self
        for y_real_value in self.y_tick_range:
            # 向上y变小，向下y变大
            y_value = self.y_middle - y_real_value * self.height_between
            self.scene.addLine(0, y_value, self.x_all, y_value, q_pen)
        return self

    def add_bin_histgram(self, data: list, q_pen_positive: QPen, q_pen_negative: QPen):
        """添加双向的直方图，正的在右边，负的在左边"""
        data.sort(reverse=True)  # 倒排
        self.height_between = self.y_middle / max(data)  # 高度间隔
        width_between = self.x_all / len(data)  # 宽度间隔
        y_tick_value_range = max(data) - min(data)
        self.y_tick_range = [i/10 for i in range(-30, 31, 1)]  # todo: 这里最好可以自动算

        for index, value in enumerate(data):  # 最大到最小
            x_pos = self.x_all - width_between * index
            q_pen = q_pen_positive if value >= 0 else q_pen_negative
            self.scene.addLine(x_pos, self.y_middle, x_pos, self.y_middle - value * self.height_between, q_pen)
        return self

    def draw(self):
        self.graph_widget.setScene(self.scene)

