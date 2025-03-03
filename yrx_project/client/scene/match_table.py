import functools
import os.path
import time

import pandas as pd
from PyQt5 import uic
from PyQt5.QtCore import QEvent, Qt, pyqtSignal
from PyQt5.QtWidgets import QApplication, QTableWidget, QMessageBox

from yrx_project.client.base import WindowWithMainWorkerBarely, BaseWorker, set_error_wrapper
from yrx_project.client.const import *
from yrx_project.client.utils.table_widget import TableWidgetWrapper
from yrx_project.scene.match_table.const import *
from yrx_project.scene.match_table.main import *
from yrx_project.utils.df_util import read_excel_file_with_multiprocessing
from yrx_project.utils.file import get_file_name_without_extension, make_zip, copy_file, open_file_or_folder_in_browser
from yrx_project.utils.iter_util import find_repeat_items
from yrx_project.utils.string_util import IGNORE_NOTHING, IGNORE_PUNC, IGNORE_CHINESE_PAREN, IGNORE_ENGLISH_PAREN
from yrx_project.utils.time_obj import TimeObj


def fill_color(main_col_indexs, matched_row_index_list, df, row_index, col_index):
    if col_index in main_col_indexs:
        if row_index in matched_row_index_list:
            return COLOR_YELLOW
    return COLOR_WHITE


def fill_color_v2(match_col_index_list, col_index):
    """
    match_col_index_list: [[], [], []]
    不同的辅助表上不同的颜色，在 COLOR_BLUE，COLOR_GREEN 中交替使用
    """
    for ind, col_index_group in enumerate(match_col_index_list):
        if ind % 2 == 0:
            if col_index in col_index_group:
                return COLOR_BLUE
        else:
            if col_index in col_index_group:
                return COLOR_GREEN
        return COLOR_WHITE


def fill_color_v3(odd_index, even_index, last_index, main_col_map, col_index, row_index):
    """
    even_index: 0，2，4 的辅助表的索引，蓝色
    odd_index: 1, 3, 5 的辅助表的索引，绿色
    last_index: 最后汇总的2个，红色
    main_col_map:
        {1: [3,5]}  # 第一列的3和5行，是匹配上的，黄色
    """
    # 列统一上色
    if col_index in even_index:
        return COLOR_BLUE
    elif col_index in odd_index:
        return COLOR_GREEN
    elif col_index in last_index:
        return COLOR_RED

    # 主表的匹配列上色
    if col_index in main_col_map:
        matched_index = main_col_map.get(col_index)
        if row_index in matched_index:
            return COLOR_YELLOW

    return COLOR_WHITE


class Worker(BaseWorker):
    custom_after_upload_signal = pyqtSignal(dict)  # 自定义信号
    custom_after_add_condition_signal = pyqtSignal(dict)  # 自定义信号
    custom_after_run_signal = pyqtSignal(dict)  # 自定义信号
    custom_view_result_signal = pyqtSignal(dict)  # 自定义信号
    custom_after_download_signal = pyqtSignal(dict)  # 自定义信号
    custom_preview_df_signal = pyqtSignal(dict)  # 自定义信号

    def my_run(self):
        stage = self.get_param("stage")  # self.equal_buffer_value.value()
        if stage == "upload":  # 任务处在上传文件的阶段
            self.refresh_signal.emit(
                f"上传文件中..."
            )
            start_upload_time = time.time()

            table_wrapper = self.get_param("table_wrapper")
            table_type = self.get_param("table_type")
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
                "table_type": table_type,
            })
        elif stage == "preview_df":
            self.refresh_signal.emit(
                f"预览表格中..."
            )
            start_preview_df_time = time.time()

            df_config = self.get_param("df_config")
            dfs = read_excel_file_with_multiprocessing([df_config])
            status_msg = f"✅预览结果成功，共耗时：{round(time.time() - start_preview_df_time, 2)}s："
            self.custom_preview_df_signal.emit({
                "df": dfs[0],
                "status_msg": status_msg
            })
        elif stage == "add_condition":  # 任务处在上传添加条件的阶段
            self.refresh_signal.emit(
                f"添加条件中..."
            )
            start_add_condition_time = time.time()

            df_main_config = self.get_param("df_main_config")
            df_help_config = self.get_param("df_help_config")
            help_tables_wrapper = self.get_param("help_tables_wrapper")
            conditions_table_wrapper = self.get_param("conditions_table_wrapper")
            table_name = help_tables_wrapper.get_data_as_df()["表名"][conditions_table_wrapper.row_length()]

            df_main_columns, df_help_columns = read_excel_file_with_multiprocessing([
                df_main_config, df_help_config
            ], only_column_name=True)

            status_msg = f"✅添加一行条件成功，共耗时：{round(time.time() - start_add_condition_time, 2)}s："
            self.custom_after_add_condition_signal.emit({
                "df_main_columns": df_main_columns,
                "df_help_columns": df_help_columns,
                "status_msg": status_msg,
                "table_name": table_name,
            })
        elif stage == "run":  # 任务处在执行的阶段
            self.refresh_signal.emit(
                f"表匹配中..."
            )
            start_run_time = time.time()

            df_main_config = self.get_param("df_main_config")
            df_help_configs = self.get_param("df_help_configs")
            conditions_df = self.get_param("conditions_df")
            result_table_wrapper = self.get_param("result_table_wrapper")
            condition_length = len(df_help_configs)

            # 构造合并条件
            match_cols_and_df = []

            df_main, *df_help_list = read_excel_file_with_multiprocessing(
                [df_main_config] + df_help_configs
            )
            read_table_time = time.time()

            # 组装match参数
            for i in range(condition_length):
                df_help = df_help_list[i]
                catch_cols = conditions_df["列：从辅助表增加"][i]
                final_catch_cols = []
                if catch_cols:
                    if isinstance(catch_cols, str):
                        final_catch_cols = [catch_cols]
                    elif isinstance(catch_cols, list):
                        final_catch_cols = catch_cols
                match_cols_and_df.append(
                    {
                        "id": conditions_df["辅助表名"][i],
                        "df": df_help,
                        "match_cols": [{
                            "main_col": conditions_df["主表匹配列"][i],
                            "match_col": conditions_df["辅助表匹配列"][i],
                        }],
                        "catch_cols": final_catch_cols,
                        "match_policy": "first",  # conditions_df["重复值策略"][i],
                        "match_ignore_policy": conditions_df["匹配忽略内容"][i],
                        "delete_policy": conditions_df["删除满足条件的行"][i],
                        "match_detail_text": conditions_df["列：匹配附加信息（文字）可编辑"][i],  # ｜ 分割的内容
                    }
                )

            # 构造是否需要额外信息
            # 1. 所有的主表匹配字段都一样
            is_all_main_col_same = len(set([conditions_df["主表匹配列"][i] for i in range(condition_length)])) == 1
            # 2. 辅助表数量大于1
            is_help_table_more_than_one = len(set([conditions_df["辅助表名"][i] for i in range(condition_length)])) > 1

            matched_df, overall_match_info, detail_match_info = match_table(
                main_df=df_main,
                match_cols_and_df=match_cols_and_df,
                add_overall_match_info=is_help_table_more_than_one
            )

            """
            {"id":{
                "time_cost": time.time() - start_for_one_df,
                "match_index_list": matched_indices,
                "unmatch_index_list": unmatched_indices,
                "no_content_index_list": no_content_indices,
                "delete_index_list": list(delete_index),

                "catch_cols": catch_cols,
                "match_extra_cols": [match_num_col_name, match_text_col_name],

                "catch_cols_index_list": [],
                "match_extra_cols_index_list": [],
                }
            }
            """
            values = detail_match_info.values()
            union_set_length, union_set_present, intersection_set_length, intersection_set_present = overall_match_info.get(
                "union_set_length"), overall_match_info.get("union_set_present"), overall_match_info.get(
                "intersection_set_length"), overall_match_info.get("intersection_set_present")
            match_table_time = time.time()

            # 填充结果表
            match_col_index_list = [v.get("catch_cols_index_list") + v.get("match_extra_cols_index_list") for v in
                                    values]
            odd_cols_index = [x for i, sublist in enumerate(match_col_index_list) for x in sublist if
                                   i % 2 != 0]  # 奇数用蓝色
            even_cols_index = [x for i, sublist in enumerate(match_col_index_list) for x in sublist if
                                    i % 2 == 0]  # 偶数用绿色
            overall_cols_index = overall_match_info.get("match_extra_cols_index_list") or []
            match_for_main_col = overall_match_info.get("match_for_main_col") or {}

            result_table_wrapper.fill_data_with_color(
                matched_df,
                cell_style_func=lambda df, row_index, col_index, odd=odd_cols_index, even=even_cols_index,
                                       last_two=overall_cols_index, match_for_main_col=match_for_main_col: fill_color_v3(
                    odd_index=odd, even_index=even, last_index=last_two, main_col_map=match_for_main_col, col_index=col_index, row_index=row_index
                )
            )
            fill_result_table = time.time()

            # 设置执行信息
            duration = round((time.time() - start_run_time), 2)
            if len(match_cols_and_df) == 1:  # 说明只有一个匹配表
                # duration = round(self.detail_match_info.values()[0].get("time_cost") * 1000, 2)
                tip = f"✅执行成功，匹配：{union_set_length}行（{union_set_present}%）"
            else:
                tip = f"✅执行成功，匹配任一条件：{union_set_length}行（{union_set_present}%）；匹配全部条件：{intersection_set_length}行（{intersection_set_present}%）"

            status_msg = \
                f"✅执行表匹配成功，共耗时：{duration}秒：读取主表+辅助表：{round(read_table_time - start_run_time, 2)}s："\
                f"表匹配：{round(match_table_time - read_table_time, 2)}s；"\
                f"填充结果表：{round(fill_result_table - match_table_time, 2)}s"

            self.custom_after_run_signal.emit({
                "tip": tip,
                "status_msg": status_msg,
                "duration": duration,
                "matched_df": matched_df,
                "overall_match_info": overall_match_info,
                "detail_match_info": detail_match_info,
                "odd_cols_index": odd_cols_index,
                "even_cols_index": even_cols_index,
                "overall_cols_index": overall_cols_index,
                "match_for_main_col": match_for_main_col,
            })
        elif stage == "view_result":
            self.refresh_signal.emit(
                f"生成预览结果..."
            )
            start_view_result = time.time()
            matched_df = self.get_param("matched_df")
            table_widget_container = self.get_param("table_widget_container")
            odd_cols_index = self.get_param("odd_cols_index")
            even_cols_index = self.get_param("even_cols_index")
            overall_cols_index = self.get_param("overall_cols_index")
            match_for_main_col = self.get_param("match_for_main_col")

            table_widget_container.fill_data_with_color(
                matched_df,
                cell_style_func=lambda df, row_index, col_index, odd=odd_cols_index, even=even_cols_index,
                                       last_two=overall_cols_index, main_col_map=match_for_main_col: fill_color_v3(
                    odd_index=odd, even_index=even, last_index=last_two, main_col_map=main_col_map, col_index=col_index, row_index=row_index
                )
            )

            duration = round((time.time() - start_view_result), 2)
            status_msg = f"✅生成放大结果成功，共耗时：{duration}秒"
            self.custom_view_result_signal.emit({
                "table_widget_wrapper": table_widget_container,
                "status_msg": status_msg,
            })

        elif stage == "download":
            self.refresh_signal.emit(
                f"合成Excel文件并下载..."
            )
            include_detail_checkbox = self.get_param("include_detail_checkbox")
            overall_match_info = self.get_param("overall_match_info")
            detail_match_info = self.get_param("detail_match_info")
            result_table_wrapper = self.get_param("result_table_wrapper")
            even_cols_index = self.get_param("even_cols_index")
            odd_cols_index = self.get_param("odd_cols_index")
            overall_cols_index = self.get_param("overall_cols_index")
            file_path = self.get_param("file_path")

            start_download = time.time()
            start_time = time.time()
            exclude_cols = []
            if not include_detail_checkbox.isChecked():  # 如果不需要详细信息，那么删除额外信息
                exclude_cols = overall_match_info.get("match_extra_cols_index_list") or []
                for i in detail_match_info.values():
                    exclude_cols.extend(i.get("match_extra_cols"))
            result_table_wrapper.save_with_color_v3(file_path, exclude_cols=exclude_cols, color_mapping={
                COLOR_BLUE.name(): even_cols_index,
                COLOR_GREEN.name(): odd_cols_index,
                COLOR_RED.name(): overall_cols_index,
                COLOR_YELLOW.name(): overall_match_info.get("match_for_main_col"),  # 是一个map key是主表匹配列的索引，value是行索引
            })
            duration = round((time.time() - start_download), 2)

            self.custom_after_download_signal.emit({
                "duration": duration,
                "status_msg": f"✅下载成功，共耗时：{duration}秒",
                "file_path": file_path,
            })


class MyTableMatchClient(WindowWithMainWorkerBarely):
    """
    重要变量
        总体
            help_info_button：点击弹出帮助信息
            release_info_button：点击弹窗版本更新信息
        第一步：添加主表、辅助表
            add_main_table_button：添加主表
            add_help_table_button：添加辅助表
            main_tables_table：主表列表
                表名 ｜ 选中工作表 ｜ 标题所在行 ｜ 操作按钮 ｜ __表路径
            help_tables_table：辅助表列表
                表名 ｜ 选中工作表 ｜ 标题所在行 ｜ 操作按钮 ｜ __表路径
        第二步：添加匹配条件
            add_condition_button：设置匹配条件
                主表匹配列 ｜ 辅助表名 ｜ 辅助表匹配列 ｜ 列：从辅助表增加 ｜ 列：匹配附加信息（文字）可编辑 ｜ 列：匹配附加信息（行数）｜操作按钮
            conditions_table：条件列表
            add_condition_help_info_button：设置匹配条件帮助信息
        第三步：执行
            run_button：执行按钮
            result_detail_text：执行详情
                 🚫执行耗时：--毫秒；共匹配：--行（--%）
            result_detail_info_button：弹出各辅助表的执行详情
            include_detail_checkbox：下载是否包括匹配附加信息
            download_result_button: 下载结果按钮
            result_table：结果表
            run_help_info_button：设置执行和下载帮助信息
    """
#     help_info_text = """
# 此场景可以用来匹配多个excel表格，解决多个表格根据同一列的关联问题，通过关联「主表」和「辅助表」实现信息的匹配，例如：
#
# 上传
# 主表：              ｜     辅助表：
#     姓名 ｜ 年龄     ｜       姓名 ｜ 性别
#     ----------     ｜       ----------
#     张三 ｜ 18      ｜        张三 ｜ 男
#     李四 ｜ 20      ｜        张三 ｜ 女
#     王五 ｜ 22      ｜        李四 ｜ 女
# ==================================================
# 结果：
# ==================================================
#     姓名 ｜ 年龄 ｜ 性别 ｜ 匹配附加信息（文字）｜ 匹配附加信息（行数）
#     ----------------------------------------------------------------------
#     张三 ｜ 18 ｜ 男    ｜     匹配到                   ｜ 2
#     李四 ｜ 20 ｜ 女    ｜     匹配到                   ｜ 1
#     王五 ｜ 22 ｜ （空） ｜    未匹配到                  ｜ 0
# """
    help_info_text = """<html>
<head>
    <title>Excel表格匹配示例</title>
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
    <h2>表格匹配场景示例</h2>
    </hr>
    <p>此场景可以用来匹配多个excel表格，解决多个表格根据同一列的关联问题，通过关联「主表」和「辅助表」实现信息的匹配，例如：</p>
    <h4>上传：主表</h4>
    <div class="table-container">
        <div class="table-wrapper1">
            <table>
                <tr>
                    <th>姓名</th>
                    <th>年龄</th>
                </tr>
                <tr>
                    <td>张三</td>
                    <td>18</td>
                </tr>
                <tr>
                    <td>李四</td>
                    <td>20</td>
                </tr>
                <tr>
                    <td>王五</td>
                    <td>22</td>
                </tr>
            </table>
        </div>
        <div class="table-wrapper1">
        <h4>上传：辅助表</h4>
            <table>
                <tr>
                    <th>姓名</th>
                    <th>性别</th>
                </tr>
                <tr>
                    <td>张三</td>
                    <td>男</td>
                </tr>
                <tr>
                    <td>张三</td>
                    <td>女</td>
                </tr>
                <tr>
                    <td>李四</td>
                    <td>女</td>
                </tr>
            </table>
        </div>
    </div>
    <h4>结果：</h4>
    <table>
        <tr>
            <th>姓名</th>
            <th>年龄</th>
            <th>性别</th>
            <th>匹配附加信息（文字）</th>
            <th>匹配附加信息（行数）</th>
        </tr>
        <tr>
            <td>张三</td>
            <td>18</td>
            <td>男</td>
            <td>匹配到</td>
            <td>2</td>
        </tr>
        <tr>
            <td>李四</td>
            <td>20</td>
            <td>女</td>
            <td>匹配到</td>
            <td>1</td>
        </tr>
        <tr>
            <td>王五</td>
            <td>22</td>
            <td>（空）</td>
            <td>未匹配到</td>
            <td>0</td>
        </tr>
    </table>
</body>
</html>"""
    release_info_text = """
v1.0.0: 实现基础版本的表匹配功能
- 主表、辅助表
- 匹配条件
- 下载结果

v1.0.1
- [修复]可能出现的单元格背景是黑色的问题

v1.0.2: 测试版
1. 支持多辅助表，多匹配条件
2. 支持设置和下载匹配附加信息
3. 增加说明按钮
4. 从辅助表携带列支持多列
5. 下载的文件用背景色区分多辅助表

v1.0.3
1. 增加全局重置功能
2. 增加状态栏显示耗时，以及错误原因
3. 所有任务异步执行，并添加loading动画
4. 下载完成后，新增：打开所在文件夹、打开文件按钮
[修复] 上传xls文件可能的报错

v1.0.4
1. 增加匹配条件选项：忽略项目
2. 添加条件时，主表选择的列给一个默认值（上一个条件的值）
3. 只要多匹配条件就会出现总评信息：任一匹配和全部匹配
[修复] xlsx文件无法选择非第一行
[修复] 全部重置按钮可能会报错
[修复] 资源文件未打包

v1.0.5
1. 优化从辅助表携带列的功能：可以补充主表而不新建列
"""

    # 第一步：上传文件的帮助信息
    step1_help_info_text = """
1. 可点击按钮或拖拽文件到表格中
2. 主表：主表的行数和最终的行数一致，主表的所有列都会在最终的结果表中
3. 辅助表：用于匹配主表
4. 调整后点击「预览」使得标题行在预览的表格的最上方
"""
    # 第二步：添加匹配条件的帮助信息
    step2_help_info_text = """
1. 添加的条件个数，不能超过辅助表的个数，且和辅助表自动一一对应
2. 列：从复制表增加：是说将匹配上列从辅助表带到主表中
3. 匹配附加信息（文字）可编辑，可以修改 匹配上｜未匹配上 ｜ 空 的文字，保证 ｜ 分割
4. 删除满足条件的行：如果选择，那么会将满足条件的行从最终的结果表中删除
"""
    # 第三步：执行与下载的帮助信息
    step3_help_info_text = """
1. 如果有多个辅助表，会增加一个综合的匹配信息
    - 任一匹配上 和 全部匹配上
    - 展示的文字和第一个匹配条件中设置的一致
2. 各辅助表统计可以查看各辅助表的匹配详情
3. 在结果表的最后面，通过颜色区分不同辅助表的匹配情况
4. 匹配到的行会在主表的匹配列中用黄色标记出来
"""

    def __init__(self):
        super(MyTableMatchClient, self).__init__()
        uic.loadUi(UI_PATH.format(file="match_table.ui"), self)  # 加载.ui文件
        self.setWindowTitle("多表匹配——By Cookie")
        self.tip_loading = self.modal(level="loading", titile="加载中...", msg=None)
        # 帮助信息
        self.help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.help_info_text, width=800, height=400))
        self.release_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.release_info_text))
        self.step1_help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.step1_help_info_text))
        self.step2_help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.step2_help_info_text))
        self.step3_help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.step3_help_info_text))
        self.demo_button.hide()  # todo 演示功能先隐藏

        # 1. 主表和辅助表的上传
        # 1.1 按钮
        self.add_main_table_button.clicked.connect(self.add_main_table)
        self.add_help_table_button.clicked.connect(self.add_help_table)
        self.reset_button.clicked.connect(self.reset_all)
        # 1.2 表格
        self.main_tables_wrapper = TableWidgetWrapper(self.main_tables_table, drag_func=self.main_drag_drop_event)  # 上传table之后展示所有table的表格
        self.help_tables_wrapper = TableWidgetWrapper(self.help_tables_table, drag_func=self.help_drag_drop_event)  # 上传table之后展示所有table的表格

        # 2. 添加匹配条件
        self.conditions_table_wrapper = TableWidgetWrapper(self.conditions_table)
        self.conditions_table_wrapper.set_col_width(3, 190).set_col_width(4, 200).set_col_width(5, 260).set_col_width(6, 150).set_col_width(7, 150)
        self.add_condition_button.clicked.connect(self.add_condition)

        # 3. 执行与下载
        self.matched_df, self.overall_match_info, self.detail_match_info = None, None, None  # 用来获取结果
        self.odd_cols_index, self.even_cols_index, self.overall_cols_index = None, None, None  # 用来标记颜色
        self.match_for_main_col = None  # 主表匹配列的映射
        self.run_button.clicked.connect(self.run)
        self.result_table_wrapper = TableWidgetWrapper(self.result_table)
        self.result_detail_info_button.clicked.connect(self.show_result_detail_info)
        # self.preview_result_button.clicked.connect(self.preview_result)
        self.download_result_button.clicked.connect(self.download_result)
        self.view_result_button.clicked.connect(self.view_result)

    def register_worker(self):
        return Worker()

    def main_drag_drop_event(self, file_names):
        if len(file_names) > 1 or len(self.main_tables_wrapper.get_data_as_df()) > 0:
            return self.modal(level="warn", msg="目前仅支持一张主表")
        self.add_table(file_names, "main")

    def help_drag_drop_event(self, file_names):
        self.add_table(file_names, "help")

    @set_error_wrapper
    def add_main_table(self, *args, **kwargs):
        if len(self.main_tables_wrapper.get_data_as_df()) > 0:
            return self.modal(level="warn", msg="目前仅支持一张主表")
        # 上传文件
        file_names = self.upload_file_modal(["Excel Files", "*.xls*"], multi=False)
        if not file_names:
            return
        self.add_table(file_names, "main")

    @set_error_wrapper
    def add_help_table(self, *args, **kwargs):
        # 上传文件
        file_names = self.upload_file_modal(["Excel Files", "*.xls*"], multi=True)
        if not file_names:
            return
        self.add_table(file_names, "help")

    @set_error_wrapper
    def reset_all(self, *args, **kwargs):
        self.main_tables_wrapper.clear()
        self.help_tables_wrapper.clear()
        self.conditions_table_wrapper.clear()
        self.result_table_wrapper.clear()
        self.statusBar.showMessage("已重置，请重新上传文件")
        self.detail_match_info = None
        self.overall_match_info = None
        self.matched_df = None
        self.result_detail_text.setText("共匹配：--行（--%）")

    # 上传文件的核心函数（调用worker）
    @set_error_wrapper
    def add_table(self, file_names, table_type):
        if isinstance(file_names, str):
            file_names = [file_names]

        for file_name in file_names:
            if not file_name.endswith(".xls") and not file_name.endswith(".xlsx"):
                return self.modal(level="warn", msg="仅支持Excel文件")

        # 根据table_type获取变量
        if table_type == "main":
            table_wrapper = self.main_tables_wrapper
        else:
            table_wrapper = self.help_tables_wrapper

        # 读取文件进行上传
        params = {
            "stage": "upload",  # 第一阶段
            "file_names": file_names,  # 上传的所有文件名
            "table_wrapper": table_wrapper,  # main_table_wrapper 或者 help_table_wrapper
            "table_type": table_type,  # main_table_wrapper 或者 help_table_wrapper
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["上传文件.", "上传文件..", "上传文件..."]).show()

    # 上传文件的后处理
    @set_error_wrapper
    def custom_after_upload(self, upload_result):
        file_names = upload_result.get("file_names")
        base_name_list = upload_result.get("base_name_list")
        sheet_names_list = upload_result.get("sheet_names_list")
        table_wrapper = upload_result.get("table_wrapper")
        status_msg = upload_result.get("status_msg")
        table_type = upload_result.get("table_type")
        for (file_name, base_name, sheet_names) in zip(file_names, base_name_list, sheet_names_list):  # 辅助表可以一次传多个，主表目前只有一个
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
                            "onclick": lambda row_index, col_index, row: self.preview_table_button(row_index, table_type=table_type),
                        }, {
                            "value": "删除",
                            "onclick": lambda row_index, col_index, row: self.delete_table_row(row_index=row_index, table_type=table_type),
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

    @set_error_wrapper
    def delete_table_row(self, row_index, table_type, *args, **kwargs):
        if table_type == "main":
            self.main_tables_wrapper.delete_row(row_index)
        else:
            table_name = self.help_tables_wrapper.get_values_by_row_index(row_index).get("表名")
            condition_df = self.conditions_table_wrapper.get_data_as_df()
            help_table_col = condition_df['辅助表名']

            # 如果没有相关条件，直接删除
            if table_name not in help_table_col.values:
                self.help_tables_wrapper.delete_row(row_index)
            # 如果在条件表中有相关条件，提示是否一并删除
            else:
                ok_or_not = self.modal(level="check_yes", msg=f"删除后，存在关联条件，是否确认删除？", default="yes")
                if ok_or_not:
                    row_index_in_condition = condition_df[help_table_col == table_name].index
                    self.conditions_table_wrapper.delete_row(row_index_in_condition[0])
                    self.help_tables_wrapper.delete_row(row_index)

    # 预览上传文件（调用worker）
    @set_error_wrapper
    def preview_table_button(self, row_index, table_type, *args, **kwargs):
        # 读取文件进行上传
        df_config = self.get_df_config_by_row_index(row_index, table_type)
        df_config["nrows"] = 10  # 实际读取的行数
        params = {
            "stage": "preview_df",  # 第一阶段
            "df_config": df_config,  # 上传的所有文件名
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
    def get_df_by_row_index(self, row_index, table_type, nrows=None, *args, **kwargs):
        df_config = self.get_df_config_by_row_index(row_index, table_type)
        df_config["nrows"] = nrows
        return read_excel_file_with_multiprocessing([df_config])[0]

    @set_error_wrapper
    def get_df_config_by_row_index(self, row_index, table_type, *args, **kwargs):
        if table_type == "main":
            table_wrapper = self.main_tables_wrapper
        else:
            table_wrapper = self.help_tables_wrapper
        path = table_wrapper.get_cell_value(row_index, 4)
        sheet_name = table_wrapper.get_cell_value(row_index, 1)  # 工作表
        row_num_for_column = table_wrapper.get_cell_value(row_index, 2)  # 列所在行
        return {
            "path": path,
            "sheet_name": sheet_name,
            "row_num_for_column": row_num_for_column,
        }

    @set_error_wrapper
    def add_condition(self, *args, **kwargs):
        """
        表选择 ｜ 主表匹配列 ｜ 辅助表匹配列 ｜ 列：从辅助表增加 ｜ 重复值策略 ｜ 匹配行标记颜色 ｜ 未匹配行标记颜色 ｜ 操作按钮
        :return:
        """
        if self.main_tables_wrapper.row_length() == 0 or self.help_tables_wrapper.row_length() == 0:
            return self.modal(level="error", msg="请先上传主表或辅助表")
        if self.conditions_table_wrapper.row_length() >= self.help_tables_wrapper.row_length():
            return self.modal(level="warn", msg="请先增加辅助表")

        df_main_config = self.get_df_config_by_row_index(0, "main")
        df_help_config = self.get_df_config_by_row_index(self.conditions_table_wrapper.row_length(), "help")

        # 读取文件进行上传
        params = {
            "stage": "add_condition",  # 第二阶段：添加条件
            "df_main_config": df_main_config,  # 主表的配置
            "df_help_config": df_help_config,  # 辅助表的配置
            "help_tables_wrapper": self.help_tables_wrapper,  # 辅助表wrapper（获取columns）
            "conditions_table_wrapper": self.conditions_table_wrapper,  # 条件table的wrapper
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["添加条件.", "添加条件..", "添加条件..."]).show()

    @set_error_wrapper
    def custom_after_add_condition(self, add_condition_result):
        status_msg = add_condition_result.get("status_msg")
        df_main_columns = add_condition_result.get("df_main_columns")
        table_name = add_condition_result.get("table_name")
        df_help_columns = add_condition_result.get("df_help_columns")

        # 获取上一个条件的主表匹配列
        default_main_col_index = None
        if self.conditions_table_wrapper.row_length() > 0:
            default_main_col = self.conditions_table_wrapper.get_cell_value(self.conditions_table_wrapper.row_length() - 1, 0)
            if default_main_col in df_main_columns:
                default_main_col_index = df_main_columns.index(default_main_col)

        # 构造级连选项
        # first_as_none = {"label": "***不从辅助表增加列***"}
        cascader_options = [{"label": NO_CATCH_COLS_OPTION}]
        for option in df_help_columns:
            column_option = {"label": option, "children": [
                {"label": ADD_COL_OPTION},
                {"label": MAKEUP_MAIN_COL, "children": [
                    {"label": main_label} for main_label in df_main_columns
                ]}
            ]}
            cascader_options.append(column_option)
        self.conditions_table_wrapper.add_rich_widget_row([
            {
                "type": "dropdown",
                "values": df_main_columns,  # 主表匹配列
                "cur_index": default_main_col_index if default_main_col_index is not None else 0,
            }, {
                "type": "readonly_text",
                "value": table_name,  # 辅助表
            }, {
                "type": "dropdown",
                "values": df_help_columns,  # 辅助表匹配列
            }, {
                "type": "dropdown",
                "values": [IGNORE_NOTHING, IGNORE_PUNC, IGNORE_CHINESE_PAREN, IGNORE_ENGLISH_PAREN],  # 重复值策略
                "cur_index": 1,  # 默认只忽略所有中英文标点符号
                "options": {
                    "multi": True,
                    "bg_colors": [COLOR_YELLOW] + [None] * 4,
                    "first_as_none": True,
                }
            # }, {
            #     "type": "dropdown",
            #     "values": ["***不从辅助表增加列***", *df_help_columns],  # 列：从辅助表增加
            #     "options": {
            #         "multi": True,
            #         "bg_colors": [COLOR_YELLOW] + [None] * len(df_help_columns),
            #         "first_as_none": True,
            #     }

            }, {
                "type": "dropdown",
                "values": cascader_options,  # 列：从辅助表增加
                "cur_index": [0],
                "options": {
                    "cascader": True,
                    # "bg_colors": [COLOR_YELLOW] + [None] * len(df_help_columns),
                    "first_as_none": True,
                }

            }, {
                "type": "editable_text",  # 列：匹配情况
                "value": " ｜ ".join(MATCH_OPTIONS),
            }, {
                "type": "readonly_text",  # 列：系统匹配到的行数
                "value": "匹配到的行数",
            }, {
                "type": "dropdown",
                "values": ["***不删除行***", *MATCH_OPTIONS],
                "options": {
                    "multi": True,
                    "first_as_none": True,
                }
            }, {
                "type": "button_group",
                "values": [
                    {
                        "value": "删除",
                        "onclick": lambda row_index, col_index, row: self.conditions_table_wrapper.delete_row(
                            row_index),
                    },
                ],

            }
        ])
        self.tip_loading.hide()
        self.set_status_text(status_msg)

    @set_error_wrapper
    def check_table_condition(self, row_index, row, *args, **kwargs):
        df_main = self.get_df_by_row_index(0, "main")
        df_help = self.get_df_by_row_index(row_index, "help")
        main_col = row["主表匹配列"]
        help_col = row["辅助表匹配列"]

        duplicate_info = check_match_table(df_main, [{
            "df": df_help,
            "match_cols": [{
                "main_col": main_col,
                "match_col": help_col,
            }],
        }])
        if duplicate_info:
            dup_values = ", ".join([str(i.get("duplicate_cols", {}).get("cell_values", [])[0]) for i in duplicate_info])
            msg = "列：{}\t重复值{}".format(help_col, dup_values)
            self.modal("warn", f"经过检查辅助表存在重复: \n{msg}")
        else:
            self.modal("info", "检查通过，辅助表不存在重复")

    @set_error_wrapper
    def run(self, *args, **kwargs):
        conditions_df = self.conditions_table_wrapper.get_data_as_df()
        if len(conditions_df) == 0:
            return self.modal(level="warn", msg="请先添加匹配条件")

        condition_length = self.conditions_table_wrapper.row_length()
        df_main_config = self.get_df_config_by_row_index(0, "main")

        # 批量读取表
        df_help_configs = [self.get_df_config_by_row_index(i, "help") for i in range(condition_length)]
        # 读取文件进行上传
        params = {
            "stage": "run",  # 表匹配
            "df_main_config": df_main_config,  # 主表的配置
            "df_help_configs": df_help_configs,  # 辅助表的配置
            "conditions_df": conditions_df,  # 条件表
            "result_table_wrapper": self.result_table_wrapper,  # 结果表的wrapper
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["表匹配.", "表匹配..", "表匹配..."]).show()

    @set_error_wrapper
    def custom_after_run(self, run_result):
        tip = run_result.get("tip")
        status_msg = run_result.get("status_msg")
        duration = run_result.get("duration")
        self.detail_match_info = run_result.get("detail_match_info")
        self.overall_match_info = run_result.get("overall_match_info")
        self.matched_df = run_result.get("matched_df")
        self.odd_cols_index = run_result.get("odd_cols_index")
        self.even_cols_index = run_result.get("even_cols_index")
        self.overall_cols_index = run_result.get("overall_cols_index")
        self.match_for_main_col = run_result.get("match_for_main_col")

        self.result_detail_text.setText(tip)
        self.tip_loading.hide()
        self.set_status_text(status_msg)
        return self.modal(level="info", msg=f"✅匹配成功，共耗时：{duration}秒")

    @set_error_wrapper
    def show_result_detail_info(self, *args, **kwargs):
        if not self.detail_match_info:
            return self.modal(level="warn", msg="请先执行")
        msg_list = []
        data = []
        for k, v in self.detail_match_info.items():
            duration = round(v.get("time_cost") * 1000, 2)
            match_percent = len(v.get('match_index_list')) / (len(v.get('match_index_list')) + len(v.get('unmatch_index_list')))
            unmatch_percent = len(v.get('unmatch_index_list')) / (len(v.get('match_index_list')) + len(v.get('unmatch_index_list')))
            delete_percent = len(v.get('delete_index_list')) / (len(v.get('match_index_list')) + len(v.get('unmatch_index_list')))
            data.append({
                "表名": k,
                # "耗时": f"{duration}s",
                "匹配行数": f"{len(v.get('match_index_list'))}（{round(match_percent * 100, 2)}%）",
                "未匹配行数": f"{len(v.get('unmatch_index_list'))}（{round(unmatch_percent * 100, 2)}%）",
                "需要删除行数": f"{len(v.get('delete_index_list'))}（{round(delete_percent * 100, 2)}%）",
            })
        self.table_modal(pd.DataFrame(data), size=(500, 200))

    @set_error_wrapper
    def view_result(self, *args, **kwargs):
        if not self.detail_match_info:
            return self.modal(level="warn", msg="请先执行")

        table_widget_container = TableWidgetWrapper()
        params = {
            "stage": "view_result",  # 阶段：预览大表格
            "matched_df": self.matched_df,  # 匹配结果
            "table_widget_container": table_widget_container,  # 匹配结果
            "odd_cols_index": self.odd_cols_index,  # 偶数辅助表相关列的索引
            "even_cols_index": self.even_cols_index,  # 奇数辅助表相关列的索引
            "overall_cols_index": self.overall_cols_index,  # 综合列的索引（最后两列）
            "match_for_main_col": self.match_for_main_col,  # 综合列的索引（最后两列）
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["生成预览结果.", "生成预览结果..", "生成预览结果..."]).show()

    @set_error_wrapper
    def custom_view_result(self, view_result):
        table_widget_wrapper = view_result.get("table_widget_wrapper")
        status_msg = view_result.get("status_msg")
        self.tip_loading.hide()
        self.set_status_text(status_msg)
        self.table_modal(
            table_widget_wrapper, size=(1200, 1000)
        )

    @set_error_wrapper
    def download_result(self, *args, **kwargs):
        if not self.detail_match_info:
            return self.modal(level="warn", msg="请先执行")
        file_path = self.download_file_modal(f"{TimeObj().time_str}_匹配结果.xlsx")
        params = {
            "stage": "download",
            "file_path": file_path,
            "include_detail_checkbox": self.include_detail_checkbox,
            "overall_match_info": self.overall_match_info,
            "detail_match_info": self.detail_match_info,
            "result_table_wrapper": self.result_table_wrapper,
            "even_cols_index": self.even_cols_index,
            "odd_cols_index": self.odd_cols_index,
            "overall_cols_index": self.overall_cols_index,

        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["合成Excel文件并下载.", "合成Excel文件并下载..", "合成Excel文件并下载..."]).show()

    @set_error_wrapper
    def custom_after_download(self, after_download_result):
        status_msg = after_download_result.get("status_msg")
        duration = after_download_result.get("duration")
        file_path = after_download_result.get("file_path")
        self.set_status_text(status_msg)
        self.tip_loading.hide()
        return self.modal(level="info", msg=f"✅下载成功，共耗时：{duration}秒", funcs=[
            # QMessageBox.ActionRole | QMessageBox.AcceptRole | QMessageBox.RejectRole
            # QMessageBox.DestructiveRole | QMessageBox.HelpRole | QMessageBox.YesRole | QMessageBox.NoRole
            # QMessageBox.ResetRole | QMessageBox.ApplyRole

            {"text": "打开所在文件夹", "func": lambda: open_file_or_folder_in_browser(os.path.dirname(file_path)), "role": QMessageBox.ActionRole},
            {"text": "打开文件", "func": lambda: open_file_or_folder_in_browser(file_path), "role": QMessageBox.ActionRole},
        ])



