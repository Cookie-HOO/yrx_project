from PyQt5 import uic
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow

from yrx_project.client.const import UI_PATH, STATIC_FILE_PATH
from yrx_project.client.scene.match_table import MyTableMatchClient


class MyClient(QMainWindow):
    def __init__(self):
        super(MyClient, self).__init__()
        uic.loadUi(UI_PATH.format(file="main.ui"), self)  # 加载.ui文件
        self.setWindowTitle("工作空间")
        self.setWindowIcon(QIcon(STATIC_FILE_PATH.format(file="app.ico")))
        self.main_tab.addTab(MyTableMatchClient(), '表匹配')
