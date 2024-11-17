import time

import pandas as pd
from PyQt5 import uic

from yrx_project.client.base import WindowWithMainWorker
from yrx_project.client.const import *
from yrx_project.client.utils.table_widget import TableWidgetWrapper
from yrx_project.scene.match_table.const import MATCH_OPTIONS
from yrx_project.scene.match_table.main import *
from yrx_project.utils.file import get_file_name_without_extension, make_zip, copy_file
from yrx_project.utils.iter_util import find_repeat_items
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


class MyTableMatchClient(WindowWithMainWorker):
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

v1.0.1: bug修复
- 修复可能出现的单元格背景是黑色的问题

v1.0.2: 功能更新
1. 支持多辅助表，多匹配条件
2. 支持设置和下载匹配附加信息
3. 增加说明按钮
4. 从辅助表携带列支持多列
5. 下载的文件用背景色区分多辅助表
"""

    # 第一步：上传文件的帮助信息
    step1_help_info_text = """"""
    # 第二步：添加匹配条件的帮助信息
    step2_help_info_text = """"""
    # 第三步：执行与下载的帮助信息
    step3_help_info_text = """"""

    def __init__(self):
        super(MyTableMatchClient, self).__init__()
        uic.loadUi(UI_PATH.format(file="match_table.ui"), self)  # 加载.ui文件
        self.setWindowTitle("表匹配——By Cookie")

        # 1. 主表和辅助表的上传
        # 1.1 按钮
        self.help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.help_info_text, width=800, height=400))
        self.release_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.release_info_text))
        self.add_main_table_button.clicked.connect(self.add_main_table)
        self.add_help_table_button.clicked.connect(self.add_help_table)
        # 1.2 表格
        self.main_tables_wrapper = TableWidgetWrapper(self.main_tables_table)  # 上传table之后展示所有table的表格
        self.help_tables_wrapper = TableWidgetWrapper(self.help_tables_table)  # 上传table之后展示所有table的表格

        # 2. 添加匹配条件
        self.conditions_table_wrapper = TableWidgetWrapper(self.conditions_table)
        self.conditions_table_wrapper.set_col_width(3, 200).set_col_width(4, 250).set_col_width(5, 150).set_col_width(6, 150)
        self.add_condition_button.clicked.connect(self.add_condition)

        # 3. 执行与下载
        self.matched_df, self.matched_index, self.unmatched_index = None, None, None
        self.run_button.clicked.connect(self.run)
        self.result_table_wrapper = TableWidgetWrapper(self.result_table)
        self.result_detail_info_button.clicked.connect(self.show_result_detail_info)
        # self.preview_result_button.clicked.connect(self.preview_result)
        self.download_result_button.clicked.connect(self.download_result)

    def add_main_table(self):
        if len(self.main_tables_wrapper.get_data_as_df()) > 0:
            return self.modal(level="warn", msg="目前仅支持一张主表")
        self.add_table("main")

    def add_help_table(self):
        self.add_table("help")

    def add_table(self, table_type):
        # 根据table_type获取变量
        if table_type == "main":
            table_wrapper = self.main_tables_wrapper
        else:
            table_wrapper = self.help_tables_wrapper
        # 上传文件
        file_names = self.upload_file_modal(["Excel Files", "*.xls*"], multi=table_type == "help")
        if not file_names:
            return
        if isinstance(file_names, str):
            file_names = [file_names]
        # 校验是否有同名文件
        base_name_list = [get_file_name_without_extension(file_name) for file_name in file_names]
        all_base_name_list = base_name_list + table_wrapper.get_data_as_df()["表名"].to_list()
        repeat_items = find_repeat_items(all_base_name_list)
        if repeat_items:
            repeat_items_str = '\n'.join(repeat_items)
            return self.modal(level="warn", msg=f"存在重复文件名，请修改后上传: \n{repeat_items_str}")
        """
        自定义表名 ｜ Excel路径 ｜ 选中工作表 ｜ 标题所在行 ｜ 标记为主表 ｜ 操作按钮
        [{
                "type": "readonly_text",
                "value": "123",
            }, {
                "type": "editable_text",
                "value": "123",
                "onchange": (row_num, col_num, row, after_change_text) => {}
            }, {
                "type": "dropdown",
                "values": ["1", "2", "3"],
                “display_values": ["1", "2", "3"]]
                "cur_value": "1"
            }, {
                "type": "global_radio",
                "value": True,
            }, {
                "type": "button_group",
                "values": [{
                   "value: "测试",
                   "onclick": (row_num, col_num, row),
                }],
            }
            ]
        """
        for index, file_name in enumerate(file_names):  # 辅助表可以一次传多个，主表目前只有一个
            base_name = get_file_name_without_extension(file_name)
            # is_main_table = bool(index == 0)
            sheet_names = pd.ExcelFile(file_name).sheet_names
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

    def delete_table_row(self, row_index, table_type):
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

    def preview_table_button(self, row_index, table_type):
        df = self.get_df_by_row_index(row_index, table_type)
        df = df.head(3)
        if len(df) >= 3:
            extra = [f'...省略{len(df)-3}行' for _ in range(df.shape[1])]
            new_row = pd.Series(extra, index=df.columns)
            df = df.append(new_row, ignore_index=True)
        self.table_modal(TableWidgetWrapper().fill_data_with_color(df).table_widget)

    def get_df_by_row_index(self, row_index, table_type):
        if table_type == "main":
            table_wrapper = self.main_tables_wrapper
        else:
            table_wrapper = self.help_tables_wrapper
        path = table_wrapper.get_cell_value(row_index, 4)
        sheet_name = table_wrapper.get_cell_value(row_index, 1)  # 工作表
        row_num_for_column = table_wrapper.get_cell_value(row_index, 2)  # 列所在行
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, header=int(row_num_for_column) - 1)
        except ValueError as e:
            df = pd.DataFrame({"error": [f"超出行数: {str(e)}"]})
        return df

    def add_condition(self):
        """
        表选择 ｜ 主表匹配列 ｜ 辅助表匹配列 ｜ 列：从辅助表增加 ｜ 重复值策略 ｜ 匹配行标记颜色 ｜ 未匹配行标记颜色 ｜ 操作按钮
        :return:
        """
        if self.main_tables_wrapper.row_length() == 0 or self.help_tables_wrapper.row_length() == 0:
            return self.modal(level="error", msg="请先上传主表或辅助表")
        if self.conditions_table_wrapper.row_length() >= self.help_tables_wrapper.row_length():
            return self.modal(level="warn", msg="请先增加辅助表")
        table_name = self.help_tables_wrapper.get_data_as_df()["表名"][self.conditions_table_wrapper.row_length()]
        df_main = self.get_df_by_row_index(0, "main")
        df_help = self.get_df_by_row_index(self.conditions_table_wrapper.row_length(), "help")

        self.conditions_table_wrapper.add_rich_widget_row([
            {
                "type": "dropdown",
                "values": df_main.columns.to_list(),  # 主表匹配列
            }, {
                "type": "readonly_text",
                "value": table_name,  # 辅助表
            }, {
                "type": "dropdown",
                "values": df_help.columns.to_list(),  # 辅助表匹配列

            # }, {
            #     "type": "dropdown",
            #     "values": ["first", "last"],  # 重复值策略
            }, {
                "type": "dropdown",
                "values": ["***不从辅助表增加列***", *df_help.columns.to_list()],  # 列：从辅助表增加
                "options": {
                    "multi": True,
                    "bg_colors": [COLOR_YELLOW] + [None] * len(df_help.columns),
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
                        "value": "检查",
                        "onclick": lambda row_index, col_index, row: self.check_table_condition(row_index, row),
                    }, {
                        "value": "删除",
                        "onclick": lambda row_index, col_index, row: self.conditions_table_wrapper.delete_row(row_index),
                    },
                ],

            }
        ])

    def check_table_condition(self, row_index, row):
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

    def run(self):
        conditions_df = self.conditions_table_wrapper.get_data_as_df()
        if len(conditions_df) == 0:
            return self.modal(level="warn", msg="请先添加匹配条件")
        start = time.time()
        df_main = self.get_df_by_row_index(0, "main")

        # 构造合并条件
        match_cols_and_df = []
        condition_length = self.conditions_table_wrapper.row_length()
        for i in range(condition_length):
            df_help = self.get_df_by_row_index(i, "help")
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
                    "delete_policy": conditions_df["删除满足条件的行"][i],
                    "match_detail_text": conditions_df["列：匹配附加信息（文字）可编辑"][i],  # ｜ 分割的内容
                }
            )

        self.matched_df, self.match_info = match_table(main_df=df_main, match_cols_and_df=match_cols_and_df)
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
        duration = round((time.time() - start) * 1000, 2)
        if len(match_cols_and_df) == 1:  # 说明只有一个匹配表
            # duration = round(self.match_info.values()[0].get("time_cost") * 1000, 2)
            match_length = len(list(self.match_info.values())[0].get("match_index_list"))
            unmatch_length = len(list(self.match_info.values())[0].get("unmatch_index_list"))
            match_present = round(match_length / (match_length + unmatch_length) * 100, 2)
            tip = f"✅执行耗时：{duration}毫秒；匹配：{match_length}行（{match_present}%）"
        else:
            tip = f"✅执行耗时：{duration}毫秒"
        self.result_detail_text.setText(tip)

        # 填充结果表
        match_col_index_list = [v.get("catch_cols_index_list") + v.get("match_extra_cols_index_list") for v in self.match_info.values()]
        # match_col_index_list = []
        # for v in self.match_info.values():
        #     match_col_index_list.append()
        self.result_table_wrapper.fill_data_with_color(
            self.matched_df,
            cell_style_func=lambda df, row_index, col_index: fill_color_v2(match_col_index_list=match_col_index_list, col_index=col_index
                                                                        ))

    def show_result_detail_info(self):
        msg_list = []
        for k, v in self.match_info.items():
            duration = round(v.get("time_cost") * 1000, 2)
            match_percent = len(v.get('match_index_list')) / (len(v.get('match_index_list') + v.get('unmatch_index_list')))
            unmatch_percent = len(v.get('unmatch_index_list')) / (len(v.get('match_index_list') + v.get('unmatch_index_list')))
            delete_percent = len(v.get('delete_percent')) / (len(v.get('match_index_list') + v.get('unmatch_index_list')))
            msg_list.append(
                f"表名：{k}\n"
                f"耗时：{duration}毫秒\n"
                f"匹配行数：{len(v.get('match_index_list'))}（{round(match_percent, 2)}）%\n"
                f"未匹配行数：{len(v.get('unmatch_index_list'))}（{round(unmatch_percent, 2)}）%\n"
                f"需要删除行数：{len(v.get('delete_index_list'))}（{round(delete_percent, 2)}）%"
            )

        self.modal(level="info", msg="==============\n".join(msg_list))

    def preview_result(self):
        self.table_modal(self.result_table_wrapper.table_widget)

    def download_result(self):
        file_path = self.download_file_modal(f"{TimeObj().time_str}_匹配结果.xlsx")

        exclude_cols = []
        if not self.include_detail_checkbox.isChecked():  # 如果不需要详细信息，那么删除额外信息
            for i in self.match_info.values():
                exclude_cols.extend(i.get("match_extra_cols"))
        self.result_table_wrapper.save_with_color(file_path, exclude_cols=exclude_cols)

