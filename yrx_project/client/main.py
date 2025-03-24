from PyQt5 import uic
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow

from yrx_project.client.const import UI_PATH, STATIC_FILE_PATH
from yrx_project.client.scene.docs_processor import MyDocsProcessorClient
from yrx_project.client.scene.match_table import MyTableMatchClient
from yrx_project.client.scene.merged_cell import MyMergedCellClient


class MyClient(QMainWindow):
    def __init__(self):
        super(MyClient, self).__init__()
        uic.loadUi(UI_PATH.format(file="main.ui"), self)  # 加载.ui文件
        self.setWindowTitle("工作空间_v1.0.5")
        self.setWindowIcon(QIcon(STATIC_FILE_PATH.format(file="app.ico")))
        self.main_tab.addTab(MyTableMatchClient(), '多表匹配')
        self.main_tab.addTab(MyDocsProcessorClient(), '文档批处理')
        self.main_tab.addTab(MyMergedCellClient(), '合并单元格')
