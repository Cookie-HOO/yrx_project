import time

from PyQt5 import uic

from yrx_project.client.base import WindowWithMainWorker
from yrx_project.client.const import *
from yrx_project.client.utils.table_widget import TableWidgetWrapper
from yrx_project.scene.match_table.main import *
from yrx_project.utils.file import get_file_name_without_extension, make_zip, copy_file
from yrx_project.utils.time_obj import TimeObj

UPLOAD_REQUIRED_FILES = ["代理期缴保费", "公司网点经营情况统计表", "农行渠道实时业绩报表"]  # 上传的文件必须要有
UPLOAD_IMPORTANT_FILE = "每日报表汇总"  # 如果有会放到tmp路径下，且会覆盖important目录中这个文件（下次使用即使不传，也是这个文件）


def fill_color(main_col_index, matched_row_index_list, df, row_index, col_index):
    if col_index == main_col_index:
        if row_index in matched_row_index_list:
            return COLOR_YELLOW


class MyTableMatchClient(WindowWithMainWorker):
    """
    重要变量
        add_table_button: 添加table按钮
        tables_table: 添加的table
            自定义表名 ｜ Excel路径 ｜ 选中工作表 ｜ 标题所在行 ｜ 标记为主表 ｜ 操作按钮
        add_condition_button：设置匹配条件
        conditions_table：
            表选择 ｜ 主表匹配列 ｜ 辅助表匹配列 ｜ 从辅助表增加列 ｜ 重复值策略 ｜ 匹配行标记颜色 ｜ 未匹配行标记颜色 ｜ 操作按钮
        run_button: 执行按钮
        result_detail_text： 🚫执行耗时：--毫秒；匹配：--行（--%）
        preview_result_button: 预览结果按钮
        download_result_button: 下载结果按钮
        result_table：结果表
    """

    help_text = """"""

    def __init__(self):
        super(MyTableMatchClient, self).__init__()
        uic.loadUi(UI_PATH.format(file="match_table.ui"), self)  # 加载.ui文件
        self.setWindowTitle("表匹配——By Cookie")

        # 1. 主表和辅助表的上传
        # 1.1 按钮
        self.help_button.clicked.connect(lambda : self.modal(level="info", msg=self.help_text))
        self.add_main_table_button.clicked.connect(self.add_main_table)
        self.add_help_table_button.clicked.connect(self.add_help_table)
        # 1.2 表格
        self.main_table_path = None
        self.main_tables_wrapper = TableWidgetWrapper(self.main_tables_table)  # 上传table之后展示所有table的表格
        self.help_table_path = None
        self.help_tables_wrapper = TableWidgetWrapper(self.help_tables_table)  # 上传table之后展示所有table的表格

        # 2. 添加匹配条件
        self.conditions_table_wrapper = TableWidgetWrapper(self.conditions_table)
        self.add_condition_button.clicked.connect(self.add_condition)

        # 3. 执行
        self.matched_df, self.matched_index, self.unmatched_index = None, None, None
        self.run_button.clicked.connect(self.run)
        self.result_table_wrapper = TableWidgetWrapper(self.result_table)
        # self.preview_result_button.clicked.connect(self.preview_result)
        self.download_result_button.clicked.connect(self.download_result)

    def add_main_table(self):
        self.add_table("main")

    def add_help_table(self):
        self.add_table("help")

    def add_table(self, table_type):
        file_names = self.upload_file_modal(["Excel Files", "*.xls*"], multi=False)
        if not file_names:
            return

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
        file_names = [file_names]
        for index, file_name in enumerate(file_names):  # 目前只有一个
            if table_type == "main":
                table_wrapper = self.main_tables_wrapper
                self.main_table_path = file_name
            else:
                table_wrapper = self.help_tables_wrapper
                self.help_table_path = file_name
            base_name = get_file_name_without_extension(file_name)
            # is_main_table = bool(index == 0)
            sheet_names = pd.ExcelFile(file_name).sheet_names
            row_num_for_columns = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]
            table_wrapper.add_rich_widget_row([
                {
                    "type": "readonly_text",  # editable_text
                    "value": base_name,
                }, {
                #     "type": "readonly_text",
                #     "value": file_name,
                # }, {
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
                            "onclick": lambda row_index, col_index, row: table_wrapper.delete_row(row_index),
                        },
                    ],

                }

            ])

    def preview_table_button(self, row_index, table_type):
        df = self.get_df_by_row_index(row_index, table_type)
        self.table_modal(TableWidgetWrapper().fill_data_with_color(df).table_widget)

    def get_df_by_row_index(self, row_index, table_type):
        if table_type == "main":
            table_wrapper = self.main_tables_wrapper
            path = self.main_table_path
        else:
            table_wrapper = self.help_tables_wrapper
            path = self.help_table_path
        sheet_name = table_wrapper.get_cell_value(row_index, 1)  # 工作表
        row_num_for_column = table_wrapper.get_cell_value(row_index, 2)  # 列所在行
        df = pd.read_excel(path, sheet_name=sheet_name, header=int(row_num_for_column) - 1)
        return df

    def add_condition(self):
        """
        表选择 ｜ 主表匹配列 ｜ 辅助表匹配列 ｜ 从辅助表增加列 ｜ 重复值策略 ｜ 匹配行标记颜色 ｜ 未匹配行标记颜色 ｜ 操作按钮
        :return:
        """
        if self.main_table_path is None or self.help_table_path is None:
            self.modal(level="error", msg="请先上传主表或辅助表")
        df_main = self.get_df_by_row_index(0, "main")
        df_help = self.get_df_by_row_index(0, "help")

        self.conditions_table_wrapper.add_rich_widget_row([
            {
                "type": "dropdown",
                "values": df_main.columns.to_list(),  # 主表匹配列
            }, {
                "type": "dropdown",
                "values": df_help.columns.to_list(),  # 辅助表匹配列
            }, {
                "type": "dropdown",
                "values": [None, *df_help.columns.to_list()],  # 从辅助表增加列
            }, {
                "type": "dropdown",
                "values": ["first", "last"],
            # }, {
            #     "type": "dropdown",
            #     "values": [None, COLOR_STR_RED, COLOR_STR_YELLOW, COLOR_STR_GREEN, COLOR_STR_BLUE],
            # }, {
            #     "type": "dropdown",
            #     "values": [None, COLOR_STR_RED, COLOR_STR_YELLOW, COLOR_STR_GREEN, COLOR_STR_BLUE],
            }, {
                "type": "button_group",
                "values": [
                    {
                        "value": "检查",
                        "onclick": lambda row_index, col_index, row: self.check_table_condition(row),
                    }, {
                        "value": "删除",
                        "onclick": lambda row_index, col_index, row: self.conditions_table_wrapper.delete_row(row_index),
                    },
                ],

            }
        ])

    def check_table_condition(self, row):
        df_main = self.get_df_by_row_index(0, "main")
        df_help = self.get_df_by_row_index(0, "help")
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
            dup_values = ", ".join([i.get("duplicate_cols", {}).get("cell_values", [])[0] for i in duplicate_info])
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
        df_help = self.get_df_by_row_index(0, "help")
        self.matched_df, self.matched_index, self.unmatched_index = match_table(main_df=df_main, match_cols_and_df=[{
            "df": df_help,
            "match_cols": [{
                "main_col": conditions_df["主表匹配列"][0],
                "match_col": conditions_df["辅助表匹配列"][0],
            }],
            "catch_cols": [conditions_df["从辅助表增加列"][0]] if conditions_df["从辅助表增加列"][0] else [],
            "match_policy": conditions_df["重复值策略"][0],
        }])
        duration = round((time.time() - start) * 1000, 2)
        match_present = round(len(self.matched_index) / (len(self.matched_index) + len(self.unmatched_index)) * 100, 2)
        tip = f"✅执行耗时：{duration}毫秒；匹配：{len(self.matched_index)}行（{match_present}%）"
        self.result_detail_text.setText(tip)

        main_col_index = df_main.columns.to_list().index(conditions_df["主表匹配列"][0])
        self.result_table_wrapper.fill_data_with_color(self.matched_df, cell_style_func=lambda df, row_index, col_index: fill_color(main_col_index=main_col_index, matched_row_index_list=self.matched_index, col_index=col_index, row_index=row_index, df=df))

    def preview_result(self):
        self.table_modal(self.result_table_wrapper.table_widget)

    def download_result(self):
        file_path = self.download_file_modal(f"{TimeObj().time_str}_匹配结果.xlsx")
        self.result_table_wrapper.save_with_color(file_path)

