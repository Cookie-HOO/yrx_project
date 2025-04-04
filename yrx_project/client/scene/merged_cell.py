import functools
import os.path
import time

import pandas as pd
from PyQt5 import uic
from PyQt5.QtCore import QEvent, Qt, pyqtSignal
from PyQt5.QtGui import QPixmap, QPainter, QBrush, QPalette
from PyQt5.QtWidgets import QApplication, QTableWidget, QMessageBox, QPushButton, QLabel, QHBoxLayout, QVBoxLayout, \
    QWidget, QMenu, QAction, QToolTip, QToolButton, QLineEdit

from yrx_project.client.base import WindowWithMainWorkerBarely, BaseWorker, set_error_wrapper
from yrx_project.client.const import *
from yrx_project.client.utils.table_widget import TableWidgetWrapper
from yrx_project.scene.match_table.const import MATCH_OPTIONS
from yrx_project.scene.match_table.main import *
from yrx_project.scene.merged_cell.const import LanguageEnum
from yrx_project.scene.merged_cell.main import do_with_code, sort_merged_cell_df
from yrx_project.utils.df_util import read_excel_file_with_multiprocessing
from yrx_project.utils.file import get_file_name_without_extension, make_zip, copy_file, open_file_or_folder
from yrx_project.utils.iter_util import find_repeat_items
from yrx_project.utils.time_obj import TimeObj


class Worker(BaseWorker):
    hide_tip_loading_signal = pyqtSignal()

    after_upload_signal = pyqtSignal(dict)  # 自定义信号
    after_do_with_code_signal = pyqtSignal(dict)  # 自定义信号
    after_download_signal = pyqtSignal(dict)  # 自定义信号

    def my_run(self):
        stage = self.get_param("stage")  # self.equal_buffer_value.value()
        if stage == "upload":
            start_upload_time = time.time()
            self.refresh_signal.emit(
                f"上传文件中..."
            )

            file_name = self.get_param("file_name")
            sheet_index = self.get_param("sheet_index")
            sheet_names_list = read_excel_file_with_multiprocessing(
                [{"path": file_name} for file_name in [file_name]],
                only_sheet_name=True, use_cache=False
            )[0]
            df = read_excel_file_with_multiprocessing(
                [{"path": file_name, "sheet_name": sheet_names_list[sheet_index], "with_merged_cells": True} for file_name in [file_name]],
                use_cache=False
            )[0]

            status_msg = \
                f"✅上传并加载第{sheet_index+1}表成功，共耗时：{round(time.time() - start_upload_time, 2)}s："

            self.after_upload_signal.emit({
                "sheet_names_list": sheet_names_list,
                "sheet_index": sheet_index,
                "df": df,
                "status_msg": status_msg
            })
        elif stage == "do_with_code":
            start_do_with_code = time.time()
            self.refresh_signal.emit(
                f"代码处理中..."
            )
            language = self.get_param("language")
            code_text = self.get_param("code_text")
            table_wrapper: TableWidgetWrapper = self.get_param("table_wrapper")
            ind = self.get_param("ind")
            res = do_with_code(language=language, code_text=code_text, df=table_wrapper.get_data_as_df(), merged_cells=table_wrapper.merged_cells, ind=ind)
            if not res["is_success"]:
                self.hide_tip_loading_signal.emit()
                return self.modal_signal.emit("error", res["error"])

            status_msg = \
                f"✅处理成功，共耗时：{round(time.time() - start_do_with_code, 2)}s："

            self.after_do_with_code_signal.emit({
                "df": res["df"],
                "table_wrapper": table_wrapper,
                "status_msg": status_msg
            })
        elif stage == "download":
            start_download_time = time.time()
            self.refresh_signal.emit(
                f"下载中..."
            )
            table_wrapper = self.get_param("table_wrapper")
            file_path = self.get_param("file_path")
            table_wrapper.save_with_merged_cells(file_path)
            status_msg = \
                f"✅下载成功，共耗时：{round(time.time() - start_download_time, 2)}s："

            self.after_download_signal.emit({
                "table_wrapper": table_wrapper,
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
- 通过python代码进行提取
- 下载功能
- 表格背景
"""

    def __init__(self):
        super(MyMergedCellClient, self).__init__()
        uic.loadUi(UI_PATH.format(file="merged_cell.ui"), self)  # 加载.ui文件
        self.setWindowTitle("合并单元格——By Cookie")
        # 帮助信息
        self.help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.help_info_text, width=800, height=400))
        self.release_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.release_info_text))
        self.demo_button.hide()  # todo 演示功能先隐藏

        self.file_name = None
        # 绑定按钮
        self.add_table_button.clicked.connect(self.add_table_button_click)

        # 主表的拖拽
        self.sheet1_table_wrapper = TableWidgetWrapper(table_widget=self.sheet1_table, drag_func=self.add_table)

        # excel的tab的切换
        self.sheet_tab.currentChanged.connect(self.sheet_tab_changed)
        # 下载文件
        self.current_table_wrapper = None
        self.all_table_wrapper = {}
        self.download_button.clicked.connect(self.download_result_button_click)

    def register_worker(self) -> BaseWorker:
        return Worker()

    def add_table_button_click(self):
        # 上传文件
        file_names = self.upload_file_modal(["Excel Files", "*.xls*"], multi=False)
        if not file_names:
            return
        self.add_table(file_names)

    def add_table(self, file_names, sheet_index=0):
        if len(file_names) != 1:
            return self.modal(level="warn", msg="只能上传一个文件")

        self.file_name = file_names[0]
        # call worker
        params = {
            "stage": "upload",  # 第一阶段
            "file_name": file_names[0],  # 上传的所有文件名
            "sheet_index": sheet_index,  # main_table_wrapper 或者 help_table_wrapper
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["上传文件.", "上传文件..", "上传文件..."]).show()

    # worker 读取文件成功后的回调，设置table表格
    def after_upload(self, after_upload_result):
        sheet_names_list = after_upload_result.get("sheet_names_list")
        df = after_upload_result.get("df")
        status_msg = after_upload_result.get("status_msg")
        sheet_index = after_upload_result.get("sheet_index")

        if self.sheet_tab.count() == 1:
            for sheet_name in sheet_names_list[1:]:
                self.sheet_tab.addTab(QWidget(), sheet_name)

        table_wrapper = self.sheet1_table_wrapper  # 默认第一个sheet的话，就不用新建table了
        if sheet_index > 0:
            table = QTableWidget()
            layout = QVBoxLayout()
            layout.addWidget(table)
            table_wrapper = TableWidgetWrapper(table_widget=table, drag_func=self.add_table)
            self.sheet_tab.widget(sheet_index).setLayout(layout)

        self.all_table_wrapper[sheet_index] = table_wrapper  # 记录所有的table_wrapper
        table_wrapper.fill_data_with_color(df, column_widget_func=self.make_title_widgets)
        path = os.path.join(STATIC_PATH, "table_bg.png")
        # table_wrapper.table_widget.setStyleSheet("QTableWidget { background-color: #FFCCCB; }")
        table_wrapper.table_widget.setStyleSheet(f"""
            QTableWidget {{
                background-image: url({path});
                background-repeat: no-repeat;
                background-position: center;
                opacity: 0.8;
            }}
        """)
        self.tip_loading.hide()
        self.set_status_text(status_msg)

    # 切换主tab之后，如果没有内容需要加载table
    def sheet_tab_changed(self, index):
        # 如果切换的tab有内容了，就不做任何操作
        cur_tab_content = self.sheet_tab.widget(index)
        if not cur_tab_content.findChildren(QWidget):
            self.add_table([self.file_name], sheet_index=index)

    def make_title_widgets(self, table_wrapper, df_columns, ind):
        # 列名
        line = QLineEdit(f"{df_columns[ind]}")
        line.setAlignment(Qt.AlignCenter)
        # 操作按钮（列名右边的）
        action_button = QToolButton(self)
        action_button.setText('...')
        action_button.setPopupMode(QToolButton.InstantPopup)
        action_button.setFixedWidth(20)
        action_button.setMenu(self.build_action_menu(table_wrapper, df_columns, ind))
        # label.setStyleSheet(f"background-color: {COLOR_BLUE.name()}")

        # 创建一个水平布局，将列名和按钮放到一起
        h_layout = QHBoxLayout()
        h_layout.addWidget(line)
        h_layout.addWidget(action_button)
        h_layout.setContentsMargins(0, 0, 8, 0)  # 左，上，右，下
        h_layout.setSpacing(0)  # 左，上，右，下

        # 创建一个小部件，将垂直布局添加到小部件中
        widget = QWidget()
        widget.setLayout(h_layout)
        widget.setStyleSheet(f"background-color: {COLOR_BLUE.name()}")

        return widget

    def build_action_menu(self, table_wrapper, df_columns, ind):
        """
        以此列筛选
        ------------
        升序（1～无穷）
        降序（无穷～1）
        自定义顺序（上传自定义序列）
        ------------
        分组
        合并（基于选定的n列）         选定的列这个按钮无法点击
        去重（对选定的n列）           未选定的列这个按钮无法点击
        ------------
        向左插入此列序号
        向左插入固定值
        向右插入此列序号
        向右插入固定值
        """
        memu_action_configs = [
            {"group": "筛选", "actions": [
                {"text": "以此列筛选", "func": self.filter_action, "tooltip": "hello"},
            ]},
            {"group": "排序", "actions": [
                {"text": "升序", "func": self.order_asc_action},
                {"text": "降序", "func": self.order_desc_action},
                {"text": "自定义顺序（需要上传序列）", "func": lambda _, table_wrapper=table_wrapper, ind=ind: self.order_custom_action(table_wrapper=table_wrapper, ind=ind)},
            ]},
            {"group": "常规", "actions": [
                {"text": "合并（基于选定的n列）", "func": self.merge_action},
                {"text": "去重（对选定的n列）", "func": self.dedup_action},
                {"text": "分组（待开发）", "func": self.group_action, "disabled": True},

            ]},
            {"group": "插入", "actions": [
                {"text": "在首列插入此列序号", "func": self.insert_num2fist_col},
                {"text": "向右插入固定值", "func": self.insert_fix2right_col},
                {"text": "向右插入提取内容", "func": lambda _, table_wrapper=table_wrapper, ind=ind: self.insert_re2right_col(table_wrapper=table_wrapper, ind=ind)},
                {"text": "向右插入自定义逻辑处理结果", "actions": [
                    {"text": "python", "func": lambda _, table_wrapper=table_wrapper, ind=ind: self.insert_custom_python2right_col(table_wrapper=table_wrapper, ind=ind)},
                ]}

            ]},
        ]

        return self.make_memu(memu_action_configs)

    def make_memu(self, menu_action_configs):
        """
        [
            {"group": "a", "actions": [
                {"text": "", "tooltip", "", "func": func, "disabled": True},
                {"text": "", "tooltip", "", "disabled": True, "actions": []},
            ]}
        ]
        """

        def _make_memu(memu, action_configs):
            for action_config in action_configs:
                action_text = action_config.get("text") or ""
                action_func = action_config.get("func") or None
                action_tooltip = action_config.get("tooltip") or ""
                action_disabled = action_config.get("disabled") or False
                sub_actions = action_config.get("actions") or []
                if sub_actions:
                    sub_menu = QMenu(action_text, self)
                    _make_memu(sub_menu, sub_actions)
                    memu.addMenu(sub_menu)
                else:
                    action = QAction(action_text, self)
                    if action_func is not None:
                        action.triggered.connect(action_func)
                    if action_tooltip:
                        action.setToolTip(action_tooltip)
                    if action_disabled:
                        action.setDisabled(True)
                    memu.addAction(action)
            memu.addSeparator()

        main_menu = QMenu(self)
        for menu_config in menu_action_configs:
            actions = menu_config.get("actions") or []
            _make_memu(main_menu, actions)
        return main_menu

    def filter_action(self):
        pass

    def order_asc_action(self):
        pass

    def order_desc_action(self):
        pass

    def order_custom_action(self, table_wrapper: TableWidgetWrapper, ind):
        df = sort_merged_cell_df(
            df=table_wrapper.get_data_as_df(),
            merged_cells=table_wrapper.merged_cells,
            col_index=ind,
            rank_order=["张三", "王五", "李四", "赵六"]
        )
        table_wrapper.fill_data_with_color(df, column_widget_func=self.make_title_widgets)


    def group_action(self):
        pass

    def merge_action(self):
        pass

    def dedup_action(self):
        pass

    def insert_num2fist_col(self):
        pass

    def insert_fix2right_col(self):
        pass

    def insert_re2right_col(self, table_wrapper: TableWidgetWrapper, ind):
        print(table_wrapper.get_data_as_df().iloc[:,ind])
        pass

    def insert_custom_python2right_col(self, table_wrapper: TableWidgetWrapper, ind):
        col_name = table_wrapper.get_cell_value(0, ind)
        init_code=f"""\
import pandas as pd

def apply(row) -> str:
    return "%" + row["{col_name}"] + "%"
    
"""
        self.code_modal(init_code=init_code, apply_func=lambda code_text: self.do_with_code(language=LanguageEnum.PYTHON, code_text=code_text, table_wrapper=table_wrapper, ind=ind))

    def do_with_code(self, language, code_text, table_wrapper: TableWidgetWrapper, ind):
        # call worker
        params = {
            "stage": "do_with_code",  # 作用代码
            "language": language,  # 代码语言
            "code_text": code_text,  # 作用的代码，里面的apply函数
            "table_wrapper": table_wrapper,  # 包装的表格的wrapper
            "ind": ind,  # 在哪一列点的
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["代码处理中.", "代码处理中..", "代码处理中..."]).show()

    def after_do_with_code(self, result):
        df = result.get("df")
        table_wrapper = result.get("table_wrapper")
        status_msg = result.get("status_msg")
        table_wrapper.fill_data_with_color(df, column_widget_func=self.make_title_widgets)

        self.tip_loading.hide()
        self.set_status_text(status_msg)

    def download_result_button_click(self):
        if not self.all_table_wrapper:
            return self.modal(level="warn", msg="请先上传文件")
        file_path = self.download_file_modal(f"{TimeObj().time_str}_合并单元格.xlsx")
        params = {
            "stage": "download",  # 下载
            "file_path": file_path,  # 下载路径
            "table_wrapper": self.all_table_wrapper[self.sheet_tab.currentIndex()],  # table组件
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["文件下载中.", "文件下载中..", "文件下载中..."]).show()

    def after_download(self, result):
        status_msg = result.get("status_msg")

        self.tip_loading.hide()
        self.set_status_text(status_msg)
