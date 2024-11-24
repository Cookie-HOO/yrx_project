import os
import time

from PyQt5.QtGui import QMovie, QIcon, QColor
from PyQt5.QtWidgets import QMessageBox, QWidget, QLabel, QVBoxLayout, QTextBrowser, QDialog
from PyQt5.QtCore import QTimer, QTime, Qt

from yrx_project.client.const import COLOR_WHITE
from yrx_project.const import STATIC_PATH



class TipWidgetWithCountDown(QMessageBox):
    """自定义的一个倒计时自动关闭的MessageBox"""
    def __init__(self, msg, count_down):
        super().__init__()
        self.time_left = count_down
        self.msg = msg
        self.initUI()
        self.exec_()

    def initUI(self):
        self.setWindowTitle(f'{self.time_left}秒后自动关闭')
        # self.setGeometry(300, 300, 250, 150)
        self.setIcon(QMessageBox.Information)
        self.setText(self.msg)
        self.setStandardButtons(QMessageBox.Ok)
        self.setDefaultButton(QMessageBox.Ok)
        # self.show()

        # 创建一个定时器
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_title)
        self.timer.start(1000)  # 每秒触发一次

    def update_title(self):
        self.time_left -= 1
        self.setWindowTitle(f'{self.time_left}秒后自动关闭')
        if self.time_left == 0:
            self.timer.stop()
            self.close()


class CustomMessageBox(QMessageBox):
    def __init__(self, title, msg, width, height, funcs, parent=None):
        super().__init__(parent)
        self.title = title
        self.msg = msg
        self.width = width
        self.height = height

        self.setWindowTitle(title)
        # self.setText(msg)

        if funcs:
            for func_item in funcs:
                text = func_item.get("text")
                func = func_item.get("func")
                # QMessageBox.ActionRole | QMessageBox.AcceptRole | QMessageBox.RejectRole
                # QMessageBox.DestructiveRole | QMessageBox.HelpRole | QMessageBox.YesRole | QMessageBox.NoRole
                # QMessageBox.ResetRole | QMessageBox.ApplyRole
                role = func_item.get("role")
                color = func_item.get("color")
                bg_color = func_item.get("bg_color")
                icon = func_item.get("icon")  # icon的绝对路径
                type_ = func_item.get("type_")  # icon的绝对路径
                if role:
                    custom_button = self.addButton(text, role)
                    if func:
                        custom_button.clicked.connect(func)
                    if color:
                        custom_button.setStyleSheet(f'QPushButton {{ color: {color.name()}; }}')
                    if bg_color:
                        custom_button.setStyleSheet(f'QPushButton {{ background-color: {bg_color.name()}; }}')
                    if icon:
                        custom_button.setIcon(QIcon(icon))

        self.setStandardButtons(QMessageBox.Ok)
        # self.setStandardButtons(QMessageBox.Ok | QMessageBox.Open | QMessageBox.Save |
        #                            QMessageBox.Cancel | QMessageBox.Close | QMessageBox.Discard |
        #                            QMessageBox.Apply | QMessageBox.Reset | QMessageBox.RestoreDefaults |
        #                            QMessageBox.Help | QMessageBox.SaveAll | QMessageBox.Yes |
        #                            QMessageBox.YesToAll | QMessageBox.No | QMessageBox.NoToAll |
        #                            QMessageBox.Abort | QMessageBox.Retry | QMessageBox.Ignore)
        # 设置自定义布局
        if "html" in self.msg:
            self.setupCustomLayout()
        else:
            self.setText(self.msg)

    def setupCustomLayout(self):
        # 获取消息框的布局
        layout = self.layout()

        # 创建一个自定义的widget和layout来控制大小
        custom_widget = QWidget()
        custom_layout = QVBoxLayout(custom_widget)

        # 添加一个自定义的label（可以添加更多的widgets）
        text_widget = QTextBrowser()
        text_widget.setHtml(self.msg)
        # label = QLabel(self.msg)
        custom_layout.addWidget(text_widget)

        # 将自定义的widget添加到消息框的布局中
        layout.addWidget(custom_widget, 0, 0)

        # 设置自定义widget的最小大小
        if self.width and self.height:
            custom_widget.setMinimumSize(self.width, self.height)


class MyQMessageBox(QWidget):
    """支持html作为msg，且支持width和height"""
    def __init__(self, title, msg, width=None, height=None, funcs=None):
        """
        funcs: [{"func": func, "text": "", "role"}]
        """
        super().__init__()
        self.title = title
        self.msg = msg
        self.width = width
        self.height = height
        self.funcs = funcs
        self.initUI()

    def initUI(self):
        # 创建并显示自定义的消息框
        msgBox = CustomMessageBox(self.title, self.msg, self.width, self.height, self.funcs)

        retval = msgBox.exec_()
        # print("返回值:", retval)


class TipWidgetWithLoading(QDialog):
    """
    tip_with_loading = TipWidgetWithLoading(title="加载中...")
    tip_with_loading.show()
    tip_with_loading.hide()
    """

    def __init__(self, titles=None):
        super().__init__()
        self.setWindowTitle("加载中...")
        self.titles = titles or []  # 多个title轮播
        self.titles_pointer = 0

        self.label = QLabel(self)
        self.label.setStyleSheet("background-color: transparent;")  # todo: 如何透明
        self.title_label = QLabel("", self)
        self.title_label.setAlignment(Qt.AlignCenter)  # 让文本居中对齐

        self.time_label = QLabel(self)
        self.time_label.setAlignment(Qt.AlignCenter)  # 让文本居中对齐

        self.movie = QMovie(os.path.join(STATIC_PATH, "loading_lwx.gif"))  # 你的GIF文件路径
        self.label.setMovie(self.movie)
        self.label.setScaledContents(True)

        self.cost_timer = QTimer(self)
        self.cost_timer.timeout.connect(self.update_cost)

        self.title_timer = QTimer(self)
        self.title_timer.timeout.connect(self.update_title)

        self.start_time = QTime()

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.addWidget(self.title_label)
        layout.addWidget(self.label)
        layout.addWidget(self.time_label)

        self.setLayout(layout)

    def showEvent(self, event):
        self.titles_pointer = 0
        self.movie.start()
        self.start_time.start()
        self.cost_timer.start(10)

        if self.titles:
            self.title_timer.start(500)
        # QTimer.singleShot(100, self.start_loading)

    def hideEvent(self, event):
        self.movie.stop()
        self.cost_timer.stop()

        self.titles_pointer = 0
        self.titles = []

    # def start_loading(self):
    #     self.titles_pointer = 0
    #     self.movie.start()
    #     self.start_time.start()
    #     self.cost_timer.start(10)
    #
    #     if self.titles:
    #         self.title_timer.start(500)

    def update_cost(self):
        elapsed = self.start_time.elapsed() / 1000.0
        self.time_label.setText("耗时：{:.2f}秒".format(elapsed))

    def update_title(self):
        if self.titles:
            title_index = self.titles_pointer % len(self.titles)
            title = self.titles[title_index]
            self.title_label.setText(title)
            self.setWindowTitle(title)
            self.titles_pointer += 1

    def set_titles(self, titles):
        self.titles = titles
        return self


# 测试代码
# from PyQt5.QtWidgets import QApplication, QDialog, QLabel, QVBoxLayout
# from PyQt5.QtGui import QMovie
# from PyQt5.QtCore import QTimer, QTime, Qt
# import sys
# app = QApplication(sys.argv)
# dialog = TipWidgetWithLoading()
# dialog.show()
# # dialog.hide()
# sys.exit(app.exec_())