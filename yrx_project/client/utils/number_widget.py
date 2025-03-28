from PyQt5.QtWidgets import QSpinBox


class CustomSpinBox(QSpinBox):
    def __init__(self, cur_value=0, step=1, prefix="", suffix="", min_num=None, max_num=None, on_change=None, parent=None):
        super().__init__(parent)
        self.setSingleStep(step or 1)  # 每次增减
        self.setValue(cur_value or 0)
        if prefix:
            self.setPrefix(prefix)   # 添加前缀
        if suffix:
            self.setSuffix(suffix)   # 添加后缀
        if min_num is not None:
            self.setMinimum(int(min_num))  # 设置最小值
        if max_num is not None:
            self.setMaximum(int(max_num))  # 设置最大值
        if on_change is not None:
            self.valueChanged.connect(on_change)  # 绑定值变化事件

        self.setStyleSheet("""
            QSpinBox {
                padding-left: 4px;  /* 左侧内边距 */
            }
        """)