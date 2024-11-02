from PyQt5.QtWidgets import QMessageBox
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
