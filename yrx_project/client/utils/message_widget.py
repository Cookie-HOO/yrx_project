from PyQt5.QtWidgets import QMessageBox, QWidget, QLabel, QVBoxLayout, QTextBrowser
from PyQt5.QtCore import QTimer


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
    def __init__(self, title, msg, width, height, parent=None):
        super().__init__(parent)
        self.title = title
        self.msg = msg
        self.width = width
        self.height = height

        self.setWindowTitle(title)
        # self.setText(msg)

        self.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)

        # 设置自定义布局
        self.setupCustomLayout()

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
        custom_widget.setMinimumSize(self.width, self.height)

class MyQMessageBox(QWidget):
    def __init__(self, title, msg, width, height):
        super().__init__()
        self.title = title
        self.msg = msg
        self.width = width
        self.height = height
        self.initUI()

    def initUI(self):
        # 创建并显示自定义的消息框
        msgBox = CustomMessageBox(self.title, self.msg, self.width, self.height)
        retval = msgBox.exec_()
        # print("返回值:", retval)

# class MyQMessageBox(QWidget):
#     def __init__(self, title, msg, width, height):
#         super().__init__()
#         self.title = title
#         self.msg = msg
#         self.width = width
#         self.height = height
#         self.initUI()
#
#     def initUI(self):
#         # 创建一个QMessageBox
#         msgBox = QMessageBox(self)
#         msgBox.setWindowTitle(self.title)
#         msgBox.setText(self.msg)
#         msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
#
#         # 设置QMessageBox的大小
#         msgBox.setMinimumSize(10000, 10000)  # 设置最小宽度和高度
#
#         # msgBox.setGeometry(1000, 1000, 10000, 10000)  # x, y, width, height
#
#         # 显示消息框
#         retval = msgBox.exec_()