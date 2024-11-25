import functools
import os.path
import time

import pandas as pd
from PyQt5 import uic
from PyQt5.QtCore import QEvent, Qt, pyqtSignal
from PyQt5.QtWidgets import QApplication, QTableWidget, QMessageBox, QPushButton, QLabel, QHBoxLayout, QVBoxLayout, \
    QWidget

from yrx_project.client.base import WindowWithMainWorkerBarely, BaseWorker, set_error_wrapper
from yrx_project.client.const import *
from yrx_project.client.utils.table_widget import TableWidgetWrapper
from yrx_project.scene.match_table.const import MATCH_OPTIONS
from yrx_project.scene.match_table.main import *
from yrx_project.utils.df_util import read_excel_file_with_multiprocessing
from yrx_project.utils.file import get_file_name_without_extension, make_zip, copy_file, open_file_or_folder_in_browser
from yrx_project.utils.iter_util import find_repeat_items
from yrx_project.utils.time_obj import TimeObj


class Worker(BaseWorker):
    after_upload_signal = pyqtSignal(dict)  # 自定义信号

    def my_run(self):
        stage = self.get_param("stage")  # self.equal_buffer_value.value()
        if stage == "upload":
            start_upload_time = time.time()
            self.refresh_signal.emit(
                f"上传文件中..."
            )

            file_name = self.get_param("file_name")
            sheet_names_list = read_excel_file_with_multiprocessing(
                [{"path": file_name} for file_name in [file_name]],
                only_sheet_name=True
            )[0]
            first_df = read_excel_file_with_multiprocessing(
                [{"path": file_name, "sheet_name": sheet_names_list[0], "with_merged_cells": True} for file_name in [file_name]],
            )[0]

            status_msg = \
                f"✅上传成功，共耗时：{round(time.time() - start_upload_time, 2)}s："

            self.after_upload_signal.emit({
                "sheet_names_list": sheet_names_list,
                "first_df": first_df,
                "status_msg": status_msg
            })


class MyMergedCellClient(WindowWithMainWorkerBarely):
    """
    重要变量
        总体
            help_info_button：点击弹出帮助信息
            release_info_button：点击弹窗版本更新信息
        上传文件
            add_table_button: 上传文件
            sheet_tab
            sheet1_table
            all_condition_button
            download_result_button
    """
    help_info_text = """"""
    release_info_text = """
v1.0.5: 
- 拖拽或点击上传表格
- 设置标题行
- 合并单元格的排序功能
"""

    def __init__(self):
        super(MyMergedCellClient, self).__init__()
        uic.loadUi(UI_PATH.format(file="merged_cell.ui"), self)  # 加载.ui文件
        self.setWindowTitle("合并单元格——By Cookie")
        self.tip_loading = self.modal(level="loading", titile="加载中...", msg=None)
        # 帮助信息
        self.help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.help_info_text, width=800, height=400))
        self.release_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.release_info_text))
        self.demo_button.hide()  # todo 演示功能先隐藏

        # 1. 绑定按钮
        self.add_table_button.clicked.connect(self.add_table_button_click)

        # 1. 主表的拖拽
        self.sheet1_table_wrapper = TableWidgetWrapper(table_widget=self.sheet1_table, drag_func=self.add_table)

    def register_worker(self) -> BaseWorker:
        return Worker()

    def add_table_button_click(self):
        # 上传文件
        file_names = self.upload_file_modal(["Excel Files", "*.xls*"], multi=False)
        if not file_names:
            return
        self.add_table(file_names)

    def add_table(self, file_names):
        if len(file_names) != 1:
            return self.modal(level="warn", msg="只能上传一个文件")

        # call worker
        params = {
            "stage": "upload",  # 第一阶段
            "file_name": file_names[0],  # 上传的所有文件名
            "table_wrapper": self,  # main_table_wrapper 或者 help_table_wrapper
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["上传文件.", "上传文件..", "上传文件..."]).show()

    def after_upload(self, after_upload_result):
        sheet_names_list = after_upload_result.get("sheet_names_list")
        first_df = after_upload_result.get("first_df")
        status_msg = after_upload_result.get("status_msg")

        # empty_row = pd.DataFrame(columns=first_df.columns)
        empty_column = pd.DataFrame(index=first_df.index)

        # 使用pd.concat函数添加空行和列
        # first_df_with_margin = pd.concat([empty_row, first_df])
        first_df_with_left_margin = pd.concat([empty_column, first_df], axis=1)
        self.sheet1_table_wrapper.fill_data_with_color(first_df, column_widget_func=self.make_title_widgets)

        self.tip_loading.hide()
        self.set_status_text(status_msg)
        print()

    def make_title_widgets(self, df_columns, ind):
        button1 = QPushButton(f"筛选")
        button2 = QPushButton(f"排序")
        label = QLabel(f"{df_columns[ind]}")
        label.setAlignment(Qt.AlignCenter)
        # label.setStyleSheet(f"background-color: {COLOR_BLUE.name()}")

        # 创建一个水平布局，将两个按钮添加到布局中
        h_layout = QHBoxLayout()
        h_layout.addWidget(button1)
        h_layout.addWidget(button2)
        h_layout.setContentsMargins(0, 0, 0, 0)  # 移除布局的边距

        # 创建一个垂直布局，将水平布局和标签添加到布局中
        v_layout = QVBoxLayout()
        v_layout.addLayout(h_layout)
        v_layout.addWidget(label)
        v_layout.setContentsMargins(0, 0, 0, 0)  # 移除布局的边距

        # 创建一个小部件，将垂直布局添加到小部件中
        widget = QWidget()
        widget.setLayout(v_layout)
        widget.setStyleSheet(f"background-color: {COLOR_BLUE.name()}")

        return widget

