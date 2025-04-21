import os
import time

import pandas as pd
from PyQt5 import uic
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QMessageBox
from pandas.core.groupby import DataFrameGroupBy

from yrx_project.client.base import WindowWithMainWorkerBarely, BaseWorker, set_error_wrapper
from yrx_project.client.const import UI_PATH
from yrx_project.client.utils.table_widget import TableWidgetWrapper
from yrx_project.scene.split_table.const import SCENE_TEMP_PATH, REORDER_COLS
from yrx_project.scene.split_table.main import SplitTable, sheets2excels, TEMP_FILE_PATH
from yrx_project.utils.df_util import read_excel_file_with_multiprocessing
from yrx_project.utils.file import get_file_name_without_extension, open_file_or_folder, copy_file
from yrx_project.utils.iter_util import find_repeat_items, dedup_list
from yrx_project.utils.time_obj import TimeObj


class Worker(BaseWorker):
    custom_after_upload_signal = pyqtSignal(dict)  # 自定义信号
    custom_preview_df_signal = pyqtSignal(dict)  # 自定义信号
    custom_after_add_split_cols_signal = pyqtSignal(dict)  # 自定义信号
    custom_init_split_table_signal = pyqtSignal(dict)  # 自定义信号

    custom_before_split_each_table_signal = pyqtSignal(dict)  # 自定义信号
    custom_after_split_each_table_signal = pyqtSignal(dict)

    custom_after_run_signal = pyqtSignal(dict)  # 自定义信号
    custom_after_sheet2excel_signal = pyqtSignal(dict)  # 自定义信号

    def my_run(self):
        stage = self.get_param("stage")  # self.equal_buffer_value.value()
        if stage == "upload":  # 任务处在上传文件的阶段
            self.refresh_signal.emit(
                f"上传文件中..."
            )
            start_upload_time = time.time()

            table_wrapper = self.get_param("table_wrapper")
            file_names = self.get_param("file_names")
            # 校验是否有同名文件
            base_name_list = [get_file_name_without_extension(file_name) for file_name in file_names]
            all_base_name_list = base_name_list + table_wrapper.get_data_as_df()["表名"].to_list()
            repeat_items = find_repeat_items(all_base_name_list)
            if repeat_items:
                repeat_items_str = '\n'.join(repeat_items)
                self.hide_tip_loading_signal.emit()
                return self.modal_signal.emit("warn", f"存在重复文件名，请修改后上传: \n{repeat_items_str}")
            check_same_name = time.time()
            sheet_names_list = read_excel_file_with_multiprocessing(
                [{"path": file_name} for file_name in file_names],
                only_sheet_name=True
            )
            read_file_time = time.time()
            status_msg = \
                f"✅上传{len(file_names)}张表成功，共耗时：{round(time.time() - start_upload_time, 2)}s："\
                f"校验文件名：{round(check_same_name - start_upload_time, 2)}s；"\
                f"读取文件：{round(read_file_time - check_same_name, 2)}s；"\

            self.custom_after_upload_signal.emit({
                "sheet_names_list": sheet_names_list,
                "table_wrapper": table_wrapper,
                "base_name_list": base_name_list,
                "file_names": file_names,
                "status_msg": status_msg,
            })
        elif stage == "preview_df":
            self.refresh_signal.emit(
                f"预览表格中..."
            )
            start_preview_df_time = time.time()

            table_wrapper = self.get_param("table_wrapper")
            row_index = self.get_param("row_index")
            path = table_wrapper.get_cell_value(row_index, 4)
            sheet_name = table_wrapper.get_cell_value(row_index, 1)  # 工作表
            row_num_for_column = table_wrapper.get_cell_value(row_index, 2)  # 列所在行
            df_config = {
                "path": path,
                "sheet_name": sheet_name,
                "row_num_for_column": row_num_for_column,
                "nrows": 10,
            }

            dfs = read_excel_file_with_multiprocessing([df_config])
            status_msg = f"✅预览结果成功，共耗时：{round(time.time() - start_preview_df_time, 2)}s："
            self.custom_preview_df_signal.emit({
                "df": dfs[0],
                "status_msg": status_msg
            })
        elif stage == "add_split_cols":  # 任务处在上传添加条件的阶段
            self.refresh_signal.emit(
                f"添加拆分列中..."
            )
            start_add_condition_time = time.time()

            table_wrapper = self.get_param("table_wrapper")
            path = table_wrapper.get_cell_value(0, 4)
            sheet_name = table_wrapper.get_cell_value(0, 1)  # 工作表
            row_num_for_column = table_wrapper.get_cell_value(0, 2)  # 列所在行
            df_config = {
                "path": path,
                "sheet_name": sheet_name,
                "row_num_for_column": row_num_for_column,
            }

            df_columns = read_excel_file_with_multiprocessing([
                df_config
            ], only_column_name=True)[0]

            status_msg = f"✅添加一行条件成功，共耗时：{round(time.time() - start_add_condition_time, 2)}s："
            self.custom_after_add_split_cols_signal.emit({
                "df_columns": df_columns,
                "status_msg": status_msg,

            })
        elif stage == "init_split_table":
            start_cal = time.time()
            table_wrapper = self.get_param("table_wrapper")
            path = table_wrapper.get_cell_value(0, 4)
            sheet_name = table_wrapper.get_cell_value(0, 1)  # 工作表
            row_num_for_column = table_wrapper.get_cell_value(0, 2)  # 列所在行
            df_config = {
                "path": path,
                "sheet_name": sheet_name,
                "row_num_for_column": row_num_for_column,
            }

            df = read_excel_file_with_multiprocessing([
                df_config
            ])[0]

            split_cols_table_wrapper = self.get_param("split_cols_table_wrapper")
            group_cols = dedup_list(split_cols_table_wrapper.get_data_as_df()["拆分列"].to_list())

            grouped = df.groupby(group_cols)
            status_msg = f"✅计算任务元信息成功，共耗时：{round(time.time() - start_cal, 2)}s："

            self.custom_init_split_table_signal.emit({
                "df": df,
                "grouped_obj": grouped,
                "group_cols": group_cols,
                "status_msg": status_msg,
            })
        elif stage == "run":  # 任务处在执行的阶段
            start_run = time.time()
            grouped_obj = self.get_param("grouped_obj")
            group_values = self.get_param("group_values")
            raw_df = self.get_param("raw_df")
            table_wrapper = self.get_param("table_wrapper")
            user_input_result = self.get_param("user_input_result")  # 用户拆分表单的结果
            path = table_wrapper.get_cell_value(0, 4)
            sheet_name = table_wrapper.get_cell_value(0, 1)  # 工作表
            row_num_for_column = table_wrapper.get_cell_value(0, 2)  # 列所在行

            names = self.get_param("names")
            total_task = len(names)

            split_table = SplitTable(path, sheet_name, row_num_for_column, raw_df, user_input_result)
            split_table.init_env()

            for index, (name, group) in enumerate(zip(names, group_values)):
                """
                group: {"col1": "a", "col2": "b", "col3": "c"}
                name: "a_b_c表"
                """
                self.refresh_signal.emit(
                     f"拆分中：{index+1}/{total_task}"
                )
                self.custom_before_split_each_table_signal.emit({
                    "row_index": index,
                })
                split_table.copy_rows_to(name, group)
                self.custom_after_split_each_table_signal.emit({
                    "row_index": index
                })
            split_table.wrap_up()
            self.refresh_signal.emit(
                f"✅执行成功，共耗时：{round(time.time() - start_run, 2)}s：",
            )
            self.custom_after_run_signal.emit({
                "duration": round(time.time() - start_run, 2),
                "split_num": total_task,
            })

        elif stage == "sheet2excel":
            sheets2excels()
            self.custom_after_sheet2excel_signal.emit({})


class MyTableSplitClient(WindowWithMainWorkerBarely):
    """
    重要变量
        总体
            help_info_button：点击弹出帮助信息
            release_info_button：点击弹窗版本更新信息
            reset_button：重置按钮
        第一步：添加主表、辅助表
            step1_help_info_button: 第一步的帮助信息
            add_table_button：添加表
            tables_table：主表列表
                表名 ｜ 选中工作表 ｜ 标题所在行 ｜ 操作按钮 ｜ __表路径
        第二步：添加添加拆分列
            step2_help_info_button: 第二步的帮助信息
            add_split_cols_button：添加拆分列
                拆分列 ｜ 操作按钮
            split_cols_table：拆分列表
        第三步：执行
            step3_help_info_button: 第三步的帮助信息
            run_button：执行按钮
            result_detail_text：执行详情
                 🚫执行耗时：--毫秒；共拆分：--个
            download_result_button: 下载结果按钮
            result_table：结果表
                拆分文件/sheet名 ｜ 行数

    """
    help_info_text = """<html>
    <head>
        <title>单表拆分示例</title>
        <style>
            table, th, td {
                border: 1px solid black;
                border-collapse: collapse;
            }
            th, td {
                padding: 10px;
                text-align: left;
            }
            .table-container {
                display: flex;
                justify-content: space-around; /* This will space the tables evenly */
                margin-bottom: 20px;
            }
            .table-wrapper {
                flex: 1; /* Each table takes equal width */
                margin: 0 10px; /* Spacing between tables */
            }
            th {
                background-color: #4CAF50; /* Green background */
                color: white; /* White text color */
                font-weight: bold; /* Bold font for headers */
            }
        </style>
    </head>
    <body>
        <h2>单表拆分示例</h2>
        </hr>
        <p>此场景可以用来将一个excel表拆分成多个excel表或sheet，需要指定一个或多个拆分列，例如按班级性别拆分：</p>
        <p>结果可以下载为单文件多个sheet，或者多个excel文件</p>
        <h4>上传：excel</h4>
        <div class="table-container">
            <div class="table-wrapper1">
                <table>
                    <tr>
                        <th>班级</th>
                        <th>学生</th>
                        <th>性别</th>
                    </tr>
                    <tr>
                        <td>一班</td>
                        <td>张三</td>
                        <td>男</td>
                    </tr>
                    <tr>
                        <td>一班</td>
                        <td>李四</td>
                        <td>男</td>
                    </tr>
                    <tr>
                        <td>二班</td>
                        <td>王五</td>
                        <td>女</td>
                    </tr>
                    <tr>
                        <td>三班</td>
                        <td>赵六</td>
                        <td>女</td>
                    </tr>
                </table>
            </div>
        </div>
        <h4>结果1：一班-男</h4>
        <table>
            <tr>
                <th>班级</th>
                <th>学生</th>
                <th>性别</th>
            </tr>
            <tr>
                <td>一班</td>
                <td>张三</td>
                <td>男</td>
            </tr>
            <tr>
                <td>一班</td>
                <td>李四</td>
                <td>男</td>
            </tr>
        </table>
        
        <h4>结果2：二班-女</h4>
        <table>
            <tr>
                <th>班级</th>
                <th>学生</th>
                <th>性别</th>
            </tr>
            <tr>
                <td>二班</td>
                <td>王五</td>
                <td>女</td>
            </tr>
        </table>
        
        <h4>结果3：三班-女</h4>
        <table>
            <tr>
                <th>班级</th>
                <th>学生</th>
                <th>性别</th>
            </tr>
            <tr>
                <td>三班</td>
                <td>赵六</td>
                <td>女</td>
            </tr>
        </table>
    </body>
    </html>"""
    release_info_text = """
v1.0.7: 
实现基础版本的表格拆分功能

v1.0.8
可由用户选择是否对预置列重排序（目前预置列只有「序号」）
"""

    # 第一步：上传文件的帮助信息
    step1_help_info_text = """
1. 可点击按钮或拖拽文件到表格中
2. 调整「标题所在行」后点击「预览」使得标题行在预览的表格的最上方
"""
    # 第二步：添加拆分列的帮助信息
    step2_help_info_text = """
1. 点击 + 从下拉列表中选择拆分列
2. 可以添加多列，拆分多列的效果，在场景说明中有例子
3. 点击删除，删除不需要的列
4. 修改了第一步上传的文件，需要删除所有列，后重新添加
"""
    # 第三步：执行与下载的帮助信息
    step3_help_info_text = """
1. 可选拆分成多个excel还是多个sheet
2. 点击执行，可以在弹窗中修改文件名/sheet名
3. 如果拆分成多个文件，下载是一个压缩包，否则是单个文件
"""

    def __init__(self):
        super(MyTableSplitClient, self).__init__()
        uic.loadUi(UI_PATH.format(file="split_table.ui"), self)  # 加载.ui文件
        self.setWindowTitle("单表拆分——By Cookie")
        self.tip_loading = self.modal(level="loading", titile="加载中...", msg=None)
        # 帮助信息
        self.help_info_button.clicked.connect(
            lambda: self.modal(level="info", msg=self.help_info_text, width=800, height=400))
        self.release_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.release_info_text))
        self.step1_help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.step1_help_info_text))
        self.step2_help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.step2_help_info_text))
        self.step3_help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.step3_help_info_text))
        self.demo_button.hide()  # todo 演示功能先隐藏

        # 1. 表格上传
        self.add_table_button.clicked.connect(self.add_table_click)
        self.reset_button.clicked.connect(self.reset_all)
        self.tables_wrapper = TableWidgetWrapper(self.tables_table, drag_func=self.main_drag_drop_event)  # 上传table之后展示所有table的表格

        # 2. 添加拆分条件
        self.add_split_cols_button.clicked.connect(self.add_split_cols)
        self.split_cols_table_wrapper = TableWidgetWrapper(self.split_cols_table)
        self.split_cols_table_wrapper.set_col_width(0, 160)

        # 3. 执行与下载
        self.run_button.clicked.connect(self.run_button_click)
        self.download_result_button.clicked.connect(self.download_result_button_click)
        self.result_table_wrapper = TableWidgetWrapper(self.result_table)
        self.result_table_wrapper.set_col_width(0, 160)

        self.done = None  # 任务执行成功的标志位，只有done了，才可以下载

    def register_worker(self):
        return Worker()

    def main_drag_drop_event(self, file_names):
        if len(file_names) > 1 or len(self.tables_wrapper.get_data_as_df()) > 0:
            return self.modal(level="warn", msg="目前仅支持一张表进行拆分")
        self.add_table(file_names)

    @set_error_wrapper
    def reset_all(self, *args, **kwargs):
        if self.done is False:
            return self.modal(level="warn", msg="正在执行中，请勿操作")
        self.done = None
        self.tables_wrapper.clear()
        self.split_cols_table_wrapper.clear()
        self.result_table_wrapper.clear()
        self.set_status_text("")
        self.result_detail_text.setText("🚫执行耗时：--毫秒；共拆分：--个")
        pass

    @set_error_wrapper
    def add_table_click(self, *args, **kwargs):
        if len(self.tables_wrapper.get_data_as_df()) > 0:
            return self.modal(level="warn", msg="目前仅支持一张表进行拆分")
        # 上传文件
        file_names = self.upload_file_modal(["Excel Files", "*.xls*"], multi=False)
        if not file_names:
            return
        self.add_table(file_names)

        # 上传文件的核心函数（调用worker）
    @set_error_wrapper
    def add_table(self, file_names):
        if isinstance(file_names, str):
            file_names = [file_names]

        for file_name in file_names:
            if not file_name.endswith(".xls") and not file_name.endswith(".xlsx"):
                return self.modal(level="warn", msg="仅支持Excel文件")

        # 读取文件进行上传
        params = {
            "stage": "upload",  # 第一阶段
            "file_names": file_names,  # 上传的所有文件名
            "table_wrapper": self.tables_wrapper,  # 用于获取上传的元信息
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["上传文件.", "上传文件..", "上传文件..."]).show()

    @set_error_wrapper
    def custom_after_upload(self, upload_result):
        file_names = upload_result.get("file_names")
        base_name_list = upload_result.get("base_name_list")
        sheet_names_list = upload_result.get("sheet_names_list")
        table_wrapper = upload_result.get("table_wrapper")
        status_msg = upload_result.get("status_msg")
        table_type = upload_result.get("table_type")
        for (file_name, base_name, sheet_names) in zip(file_names, base_name_list,
                                                       sheet_names_list):  # 辅助表可以一次传多个，主表目前只有一个
            row_num_for_columns = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]
            table_wrapper.add_rich_widget_row([
                {
                    "type": "readonly_text",  # editable_text
                    "value": base_name,
                }, {
                    "type": "dropdown",
                    "values": sheet_names,
                    "cur_index": 0,
                }, {
                    "type": "dropdown",
                    "values": row_num_for_columns,
                    "cur_index": 0,
                    # }, {
                    #     "type": "global_radio",
                    #     "value": is_main_table,
                }, {
                    "type": "button_group",
                    "values": [
                        {
                            "value": "预览",
                            "onclick": lambda row_index, col_index, row: self.preview_table_button_click(row_index,
                                                                                                   table_type=table_type),
                        }, {
                            "value": "删除",
                            "onclick": lambda row_index, col_index, row: self.tables_wrapper.delete_row(row_index),
                            # "onclick": lambda row_index, col_index, row: self.help_tables_wrapper.delete_row(row_index),
                        },
                    ],

                }, {
                    "type": "readonly_text",
                    "value": file_name,
                },

            ])

        self.tip_loading.hide()
        self.set_status_text(status_msg)

    # 预览上传文件（调用worker）
    @set_error_wrapper
    def preview_table_button_click(self, row_index, *args, **kwargs):
        # 读取文件进行上传
        params = {
            "stage": "preview_df",  # 第一阶段
            "table_wrapper": self.tables_wrapper,
            "row_index": row_index,
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["预览表格.", "预览表格..", "预览表格..."]).show()

    @set_error_wrapper
    def custom_preview_df(self, preview_result):
        df = preview_result.get("df")
        status_msg = preview_result.get("status_msg")
        max_rows_to_show = 10
        if len(df) >= max_rows_to_show:
            extra = [f'...省略剩余行' for _ in range(df.shape[1])]
            new_row = pd.Series(extra, index=df.columns)
            # 截取前 max_rows_to_show 行，再拼接省略行信息
            df = pd.concat([df.head(max_rows_to_show), pd.DataFrame([new_row])], ignore_index=True)
        self.tip_loading.hide()
        self.set_status_text(status_msg)
        self.table_modal(df, size=(400, 200))

    @set_error_wrapper
    def add_split_cols(self, *args, **kwargs):
        """
        拆分列 ｜ 操作按钮
        :return:
        """
        if self.done is False:
            return self.modal(level="warn", msg="正在执行中，请勿操作")
        if self.tables_wrapper.row_length() == 0:
            return self.modal(level="error", msg="请先上传待拆分表")

        # 读取文件进行上传
        params = {
            "stage": "add_split_cols",  # 第二阶段：添加条件
            "table_wrapper": self.tables_wrapper,
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["添加拆分列.", "添加拆分列..", "添加拆分列..."]).show()

    @set_error_wrapper
    def custom_after_add_split_cols(self, add_split_cols_result):
        status_msg = add_split_cols_result.get("status_msg")
        df_columns = add_split_cols_result.get("df_columns")

        self.split_cols_table_wrapper.add_rich_widget_row([
            {
                "type": "dropdown",
                "values": df_columns,  # 主表匹配列
                "cur_index": 0,
            },{
                "type": "button_group",
                "values": [
                    {
                        "value": "删除",
                        "onclick": lambda row_index, col_index, row: self.split_cols_table_wrapper.delete_row(
                            row_index),
                    },
                ],
            }
        ])
        self.tip_loading.hide()
        self.set_status_text(status_msg)

    @set_error_wrapper
    def run_button_click(self, *args, **kwargs):
        """
        1. 点击执行后，是一个弹窗，可以配置名字，且显示个数，比如
            共 37 个：{院系}-{教师}
        2. 执行时一律拆分成sheet
            下载时根据 拆成多个excel文件，还是多个sheet决定下载成什么
        """
        if self.tables_wrapper.row_length() == 0 or self.split_cols_table_wrapper.row_length() == 0:
            return self.modal(level="warn", msg="请先上传文件和指定拆分列")
        if self.done is False:
            return self.modal(level="warn", msg="正在执行中，请勿操作")

        # 读取文件进行上传
        params = {
            "stage": "init_split_table",  # 第二阶段：添加条件
            "table_wrapper": self.tables_wrapper,
            "split_cols_table_wrapper": self.split_cols_table_wrapper,
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["任务元信息.", "任务元信息..", "任务元信息..."]).show()

    @set_error_wrapper
    def custom_init_split_table(self, init_split_table_result):
        """
        1. 计算元信息后的回调，注入要拆分的个数，进行弹窗提示
        2. 初始化结果表
        3. 开始执行任务
        """
        status_msg = init_split_table_result.get("status_msg")
        self.set_status_text(status_msg)

        group_cols = init_split_table_result.get("group_cols")
        grouped_obj: DataFrameGroupBy = init_split_table_result.get("grouped_obj")
        raw_df: pd.DataFrame = init_split_table_result.get("df")
        split_num = grouped_obj.ngroups
        self.tip_loading.hide()

        # 格式
        cols = self.split_cols_table_wrapper.get_data_as_df()["拆分列"].to_list()
        default_split_name_format = "-".join(["{" + col + "}"for col in cols])

        # 个数
        need_split, result = self.modal(level="form", msg=f"确定要拆分吗，即将拆分 {split_num} 个？", fields_config = [
            *[{
                "id": f"reorder_{reorder_col}",
                "type": "checkbox",
                "label": f"在拆分结果中，重排序「{reorder_col}」列",
                "default": True,
                "show_if": reorder_col in raw_df.columns,
            } for reorder_col in REORDER_COLS],
            {
                "id": "split_name_format",
                "type": "editable_text",
                "label": "拆分文件/sheet名格式",
                "default": default_split_name_format,
                "placeholder": "文件/sheet名格式",
                "limit": lambda x: "格式不能为空" if len(x) == 0 else "",
            },
        ])
        if not need_split:
            return
        self.done = False  # 开始进入计算周期
        split_name_format = result.get("split_name_format")
        # 获取分组统计结果
        size_series = grouped_obj.size().reset_index(name='行数')
        # 生成格式化分组名称
        size_series['拆分文件/sheet'] = size_series.apply(
            lambda row: split_name_format.format(**{col: row[col] for col in group_cols}),
            axis=1
        )
        df = size_series[['拆分文件/sheet', '行数']]
        df["拆分文件/sheet"] = df["拆分文件/sheet"].apply(lambda x: x.replace("%", "_").replace("/", "_").replace("\\", "_").replace("?", "_").replace("*", "_").replace("[", "_").replace("]", "_").replace(":", "_").replace("：", "_").replace("'", "_"))
        df["拆分文件/sheet"] = df["拆分文件/sheet"].apply(lambda x: x[:20] if len(x) > 20 else x)
        df["拆分文件/sheet"] = df["拆分文件/sheet"].apply(lambda x: x if x else "%EMPTY%")

        # 初始化结果表
        self.result_table_wrapper.fill_data_with_color(df)

        # 开始执行任务
        groups = []  # 存储结果的列表

        # 遍历所有分组键
        for key in grouped_obj.groups.keys():
            # 将元组键转换为字典
            if not isinstance(key, tuple):
                key = (key,)
            group_dict = {}
            for col_name, value in zip(group_cols, key):
                group_dict[col_name] = value
            groups.append(group_dict)
        """
        [{"col1": "一班", "col2": "男"}, {"col1": "二班", "col2": "女"}]
        """

        params = {
            "stage": "run",  # 第三阶段：执行
            "table_wrapper": self.tables_wrapper,
            "grouped_obj": grouped_obj,
            "group_values": groups,
            "raw_df": raw_df,
            "user_input_result": result,
            "names": df["拆分文件/sheet"].to_list(),
        }
        self.worker.add_params(params).start()
        # self.tip_loading.set_titles(["表拆分.", "表拆分..", "表拆分..."]).show()

    @set_error_wrapper
    def custom_before_split_each_table(self, before_split_table_result):
        """准备要拆分的那个table时进行回调
        修改对应结果表的行头emoji
        result_table：结果表
            拆分文件/sheet名 ｜ 行数
        """
        row_index = before_split_table_result.get("row_index")
        self.result_table_wrapper.update_vertical_header(row_index, "🏃")
        pass

    @set_error_wrapper
    def custom_after_split_each_table(self, after_split_table_result):
        """拆分完的那个table时进行回调
        修改对应结果表的行头emoji
        result_table：结果表
            拆分文件/sheet名 ｜ 行数
        """
        row_index = after_split_table_result.get("row_index")
        self.result_table_wrapper.update_vertical_header(row_index, "✅")
        pass

    @set_error_wrapper
    def custom_after_run(self, after_run_result):
        """
        拆分任务结束的回调
        result_detail_text：执行详情
             🚫执行耗时：--毫秒；共拆分：--个
        """
        status_msg = after_run_result.get("status_msg")
        duration = after_run_result.get("duration")
        split_num = after_run_result.get("split_num")
        self.set_status_text(status_msg)
        self.done = True
        msg = f"✅执行耗时：{duration}秒；共拆分：{split_num}个"
        self.result_detail_text.setText(msg)
        self.modal(level="info", msg=msg + "\n可以通过「下载结果」按钮下载拆分结果")
        pass

    @set_error_wrapper
    def download_result_button_click(self, *args, **kwargs):
        """
        split2excel_radio：拆分成多个excel的radio
        split2sheet_radio：拆分成多个sheet的radio
        """
        if not self.done:
            return self.modal(level="warn", msg="任务没有执行完成，无法下载")
        need_download, result = self.modal(level="form", msg=f"下载结果", fields_config=[
            {
                "id": "download_format",
                "type": "radio_group",
                "labels": ["拆分成多个excel文件", "拆分成单文件多sheet"],
                "default": "拆分成多个excel文件",
            },
        ])
        if not need_download:
            return

        # 拆分sheet，直接下载
        if result.get("download_format") == "拆分成单文件多sheet":
            file_path = self.download_file_modal(f"{TimeObj().time_str}_拆分结果.xlsx")
            copy_file(TEMP_FILE_PATH, file_path)
            return self.modal(level="info", msg=f"✅下载成功", funcs=[
                {"text": "打开所在文件夹", "func": lambda: open_file_or_folder(os.path.dirname(file_path)),
                 "role": QMessageBox.ActionRole},
                {"text": "打开文件", "func": lambda: open_file_or_folder(file_path), "role": QMessageBox.ActionRole},
            ])
        # 拆分excel，需要异步将sheet转成excel
        params = {
            "stage": "sheet2excel",  # 第三阶段：执行
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["准备下载文件，对齐格式.", "准备下载文件，对齐格式..", "准备下载文件，对齐格式..."]).show()

    @set_error_wrapper
    def custom_after_sheet2excel(self, after_download_result):
        status_msg = after_download_result.get("status_msg")
        self.set_status_text(status_msg)
        self.tip_loading.hide()
        duration = after_download_result.get("duration")
        is_success, file_path = self.download_zip_from_path(SCENE_TEMP_PATH, "拆分结果")
        if is_success:
            return self.modal(level="info", msg=f"✅下载压缩包成功，共耗时：{duration}秒", funcs=[
                {"text": "打开所在文件夹", "func": lambda: open_file_or_folder(os.path.dirname(file_path)),
                 "role": QMessageBox.ActionRole},
                {"text": "打开文件", "func": lambda: open_file_or_folder(file_path), "role": QMessageBox.ActionRole},
            ])
