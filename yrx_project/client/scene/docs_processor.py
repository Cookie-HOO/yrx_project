import os
import time
import typing

from PyQt5 import uic
from PyQt5.QtCore import pyqtSignal

from yrx_project.client.base import WindowWithMainWorkerBarely, BaseWorker, set_error_wrapper
from yrx_project.client.const import UI_PATH
from yrx_project.client.utils.button_menu_widget import ButtonMenuWrapper
from yrx_project.client.utils.table_widget import TableWidgetWrapper
from yrx_project.client.utils.tree_file_widget import TreeFileWrapper
from yrx_project.const import PROJECT_PATH, TEMP_PATH
from yrx_project.scene.docs_processor.base import ActionContext
from yrx_project.scene.docs_processor.const import ACTION_MAPPING
from yrx_project.client.scene.docs_processor_adapter import run_with_actions, build_action_types_menu
from yrx_project.utils.file import get_file_name_without_extension
from yrx_project.utils.iter_util import find_repeat_items
from yrx_project.utils.time_obj import TimeObj


class Worker(BaseWorker):
    custom_after_upload_signal = pyqtSignal(dict)  # 自定义信号
    # custom_after_add_condition_signal = pyqtSignal(dict)  # 自定义信号
    custom_after_run_signal = pyqtSignal(dict)  # 自定义信号
    custom_update_progress_signal = pyqtSignal(float)
    # custom_view_result_signal = pyqtSignal(dict)  # 自定义信号
    # custom_after_download_signal = pyqtSignal(dict)  # 自定义信号
    # custom_preview_df_signal = pyqtSignal(dict)  # 自定义信号

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
            all_base_name_list = base_name_list + table_wrapper.get_data_as_df()["文档名称"].to_list()
            repeat_items = find_repeat_items(all_base_name_list)
            if repeat_items:
                repeat_items_str = '\n'.join(repeat_items)
                self.hide_tip_loading_signal.emit()
                return self.modal_signal.emit("warn", f"存在重复文件名，请修改后上传: \n{repeat_items_str}")

            check_same_name = time.time()
            # pages = get_docx_pages_with_multiprocessing(file_names)
            # read_file_time = time.time()
            status_msg = \
                f"✅上传{len(file_names)}张表成功，共耗时：{round(time.time() - start_upload_time, 2)}s："\
                f"校验文件名：{round(check_same_name - start_upload_time, 2)}s；"\
                # f"读取文件：{round(read_file_time - check_same_name, 2)}s；"\

            self.custom_after_upload_signal.emit({
                # "pages": pages,
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

            df_config = self.get_param("df_config")
            dfs = read_excel_file_with_multiprocessing([df_config])
            status_msg = f"✅预览结果成功，共耗时：{round(time.time() - start_preview_df_time, 2)}s："
            self.custom_preview_df_signal.emit({
                "df": dfs[0],
                "status_msg": status_msg
            })

        elif stage == "run":  # 任务处在执行的阶段
            self.refresh_signal.emit(
                f"文档处理中..."
            )
            start_run_time = time.time()

            """
            "stage": "run",  # run
            "df_docs": df_docs,  # 文档的路径
            "df_actions": df_actions,  # 动作流
            """

            df_docs = self.get_param("df_docs")
            df_actions = self.get_param("df_actions")

            def callback(ctx: ActionContext):
                file_name = "--"
                if ctx.file_path:
                    file_name = get_file_name_without_extension(ctx.file_path)
                self.refresh_signal.emit(f"文档处理中...阶段: {ctx.command_container.step_and_name} 进度：{ctx.done_task_num}/{ctx.total_task_num}; 文件: {file_name}: 操作: {ctx.command.action_name}")
                self.custom_update_progress_signal.emit(ctx.done_task_num / ctx.total_task_num)
            run_with_actions(
                input_paths=df_docs["__文档路径"].to_list(),
                df_actions=df_actions,
                after_each_action_func=callback,
            )

            # 设置执行信息
            duration = round((time.time() - start_run_time), 2)
            tip = f"✅执行成功"

            status_msg = \
                f"✅批量文档处理成功，共耗时：{duration}秒"

            self.custom_after_run_signal.emit({
                "tip": tip,
                "status_msg": status_msg,
                "duration": duration,
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


class MyDocsProcessorClient(WindowWithMainWorkerBarely):
    """
        重要变量
            总体
                help_info_button：点击弹出帮助信息
                release_info_button：点击弹窗版本更新信息
                reset_button：重置所有
            第一步：docs
                add_docs_button：添加word文档
                docs_table
                    文档名称 | 页数 | 操作按钮 | __文档路径
            第二步：定义动作流
                add_actions_combo：添加的操作类型：定位、选择、修改、合并
                add_actions_button：设置匹配条件
                actions_table：动作流表格
                    顺序 ｜ 类型 ｜ 动作 ｜ 动作内容 | 操作按钮
            第三步：执行
                run_help_info_button：设置执行和下载帮助信息
                run_button：执行按钮
                result_detail_text：执行详情
                     🚫执行耗时：--毫秒；共匹配：--行（--%）
                result_tree：结果文件的树状结构
                run_progress_bar：进度条
                download_result_button: 下载结果按钮
                result_preview_grid_layout：结果文件的预览
                    test1_preview_img QLabel 测试缩略图
                    test2_preview_img QLabel 测试缩略图
                    test3_preview_img QLabel 测试缩略图
                    test4_preview_img QLabel 测试缩略图
                result_preview_col_name_text
                preview_col_num_add_button
                preview_col_num_sub_button
        """

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
        <h2>多文档操作场景</h2>
        </hr>
        <p>此场景可以用来操作多个word文档，定义执行操作流，例如：</p>
        <h4>上传：通过拖拽或点击上传文档后得到列表</h4>
        <div class="table-container">
            <div class="table-wrapper1">
                <table>
                    <tr>
                        <th>文档名称</th>
                        <th>页数</th>
                        <th>操作按钮</th>
                    </tr>
                    <tr>
                        <td>第1篇文档</td>
                        <td>2页</td>
                        <td>|删除|</td>
                    </tr>
                    <tr>
                        <td>第2篇文档</td>
                        <td>3页</td>
                        <td>|删除|</td>
                    </tr>
                    <tr>
                        <td>第3篇文档</td>
                        <td>2页</td>
                        <td>|删除|</td>
                    </tr>
                </table>
            </div>
            <div class="table-wrapper1">
            <h4>定义：动作流</h4>
                <table>
                    <tr>
                        <th>顺序</th>
                        <th>类型</th>
                        <th>动作</th>
                        <th>动作内容</th>
                        <th>操作按钮</th>
                    </tr>
                    <tr>
                        <td>1</td>
                        <td>定位</td>
                        <td>搜索</td>
                        <td>=</td>
                        <td>|向上|向下|删除|</td>
                    </tr>
                    <tr>
                        <td>2</td>
                        <td>定位</td>
                        <td>向左移动</td>
                        <td> 1 </td>
                        <td>|向上|向下|删除|</td>
                    </tr>
                    <tr>
                        <td>3</td>
                        <td>选择</td>
                        <td>选择当前单元格</td>
                        <td> --- </td>
                        <td>|向上|向下|删除|</td>
                    </tr>
                    <tr>
                        <td>4</td>
                        <td>修改</td>
                        <td>文字替换</td>
                        <td> abc </td>
                        <td>|向上|向下|删除|</td>
                    </tr>
                    <tr>
                        <td>5</td>
                        <td>总体</td>
                        <td>合并成一个文档</td>
                        <td> -- </td>
                        <td>|向上|向下|删除|</td>
                    </tr>
                </table>
            </div>
        </div>
        <h4>结果：</h4>
    </body>
    </html>"""
    release_info_text = """
    v1.0.6: 实现基础版本的文档聚合
    - 上传多个文档
    - 实现
        - 定位：搜索、移动
        - 选择：选择当前单元格内容
        - 修改：替换
        - 聚合：合并
    - 下载结果
    """

    # 第一步：上传文件的帮助信息
    step1_help_info_text = """
    1. 可点击按钮或拖拽文档到表格中：目前只支持docx格式
    2. 可点击预览查看上传的文档 TODO
    """
    # 第二步：添加动作流的帮助信息
    step2_help_info_text = """
    1. 点击添加，会显示添加的动作类型，目前支持：定位、选择、修改、聚合
    2. 指定动作类型后，在动作中选择一个对应的动作
    3. 输入动作内容
    4. 操作按钮中可以：向上移动、向下移动、删除
    """
    # 第三步：执行与下载的帮助信息
    step3_help_info_text = """
    1. 第二步输入的动作内容可能存在问题，执行后，会进行提示
    2. 预览的结果可能存在格式的问题，仅作示意，以下载的内容为准 TODO
    """

    def __init__(self):
        super(MyDocsProcessorClient, self).__init__()
        uic.loadUi(UI_PATH.format(file="process_docs.ui"), self)  # 加载.ui文件
        self.setWindowTitle("文档批处理——By Cookie")
        self.tip_loading = self.modal(level="loading", titile="加载中...", msg=None)
        # 帮助信息
        self.help_info_button.clicked.connect(
            lambda: self.modal(level="info", msg=self.help_info_text, width=800, height=400))
        self.release_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.release_info_text))
        self.step1_help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.step1_help_info_text))
        self.step2_help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.step2_help_info_text))
        self.step3_help_info_button.clicked.connect(lambda: self.modal(level="info", msg=self.step3_help_info_text))
        self.demo_button.hide()  # todo 演示功能先隐藏

        # 1. 批量上传文档
        # 1.1 按钮
        self.add_docs_button.clicked.connect(self.add_docs)
        # self.reset_button.clicked.connect(self.reset_all)
        # 1.2 表格
        self.docs_tables_wrapper = TableWidgetWrapper(self.docs_table,
                                                      drag_func=self.docs_drag_drop_event)  # 上传docs之后展示所有table的表格
        # self.help_tables_wrapper = TableWidgetWrapper(self.help_tables_table,
        #                                               drag_func=self.help_drag_drop_event)  # 上传table之后展示所有table的表格
        #
        # # 2. 添加动作流
        self.actions_table_wrapper = TableWidgetWrapper(self.actions_table).set_col_width(1, 150).set_col_width(4, 200)
        self.add_action_button_menu = ButtonMenuWrapper(self, self.add_action_button, build_action_types_menu(self.actions_table_wrapper))

        # self.add_action_button.clicked.connect(self.add_action)
        #
        # # 3. 执行与下载
        # self.matched_df, self.overall_match_info, self.detail_match_info = None, None, None  # 用来获取结果
        # self.odd_cols_index, self.even_cols_index, self.overall_cols_index = None, None, None  # 用来标记颜色
        # self.match_for_main_col = None  # 主表匹配列的映射
        self.run_button.clicked.connect(self.run)
        self.run_progress_bar.setValue(0)
        self.tree_file_wrapper = TreeFileWrapper(self.result_tree, TEMP_PATH)


        """                result_preview_col_num_text
                preview_col_num_add_button
                preview_col_num_sub_button"""
        self.preview_col_num_add_button.clicked.connect(lambda: self.update_preview_col_num(1))
        self.preview_col_num_sub_button.clicked.connect(lambda: self.update_preview_col_num(-1))


        # self.worker.custom_after_upload_signal.connect(self.custom_after_upload)
        # self.result_table_wrapper = TableWidgetWrapper(self.result_table)
        # self.result_detail_info_button.clicked.connect(self.show_result_detail_info)
        # # self.preview_result_button.clicked.connect(self.preview_result)
        # self.download_result_button.clicked.connect(self.download_result)
        # self.view_result_button.clicked.connect(self.view_result)

    def register_worker(self):
        return Worker()

    def main_drag_drop_event(self, file_names):
        if len(file_names) > 1 or len(self.main_tables_wrapper.get_data_as_df()) > 0:
            return self.modal(level="warn", msg="目前仅支持一张主表")
        self.add_table(file_names, "main")

    def docs_drag_drop_event(self, file_names):
        self.add_doc(file_names)

    # @set_error_wrapper
    # def add_main_table(self, *args, **kwargs):
    #     if len(self.main_tables_wrapper.get_data_as_df()) > 0:
    #         return self.modal(level="warn", msg="目前仅支持一张主表")
    #     # 上传文件
    #     file_names = self.upload_file_modal(["Excel Files", "*.xls*"], multi=False)
    #     if not file_names:
    #         return
    #     self.add_table(file_names, "main")
    #
    @set_error_wrapper
    def add_docs(self, *args, **kwargs):
        # 上传文件
        file_names = self.upload_file_modal(["Word Files", "*.docx"], multi=True)
        if not file_names:
            return
        self.add_doc(file_names)
    #
    # @set_error_wrapper
    # def reset_all(self, *args, **kwargs):
    #     self.main_tables_wrapper.clear()
    #     self.help_tables_wrapper.clear()
    #     self.conditions_table_wrapper.clear()
    #     self.result_table_wrapper.clear()
    #     self.statusBar.showMessage("已重置，请重新上传文件")
    #     self.detail_match_info = None
    #     self.overall_match_info = None
    #     self.matched_df = None
    #     self.result_detail_text.setText("共匹配：--行（--%）")
    #
    # 上传文件的核心函数（调用worker）
    @set_error_wrapper
    def add_doc(self, file_names):
        if isinstance(file_names, str):
            file_names = [file_names]

        for file_name in file_names:
            if not file_name.endswith(".docx"):
                return self.modal(level="warn", msg="仅支持docx文件")

        table_wrapper = self.docs_tables_wrapper

        # 读取文件进行上传
        params = {
            "stage": "upload",  # 第一阶段
            "file_names": file_names,  # 上传的所有文件名
            "table_wrapper": table_wrapper,  # main_table_wrapper 或者 help_table_wrapper
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["上传文件.", "上传文件..", "上传文件..."]).show()

    # 上传文件的后处理
    @set_error_wrapper
    def custom_after_upload(self, upload_result):
        # pages = upload_result.get("pages")
        file_names = upload_result.get("file_names")
        base_name_list = upload_result.get("base_name_list")
        table_wrapper = upload_result.get("table_wrapper")
        status_msg = upload_result.get("status_msg")
        for (file_name, base_name) in zip(file_names, base_name_list):  # 辅助表可以一次传多个，主表目前只有一个
            table_wrapper.add_rich_widget_row([
                {
                    "type": "readonly_text",  # editable_text
                    "value": base_name,
                }, {
                    "type": "readonly_text",  # editable_text
                    "value": "-1",
                # }, {
                #     "type": "dropdown",
                #     "values": sheet_names,
                #     "cur_index": 0,
                # }, {
                #     "type": "dropdown",
                #     "values": row_num_for_columns,
                #     "cur_index": 0,
                    # }, {
                    #     "type": "global_radio",
                    #     "value": is_main_table,
                }, {
                    "type": "button_group",
                    "values": [
                        # {
                        #     "value": "预览",
                        #     "onclick": lambda row_index, col_index, row: self.preview_table_button(row_index,
                        #                                                                            table_type=table_type),
                        # },
                        {
                            "value": "删除",
                            "onclick": lambda row_index, col_index, row: self.delete_table_row(row_index=row_index,
                                                                                               table_type="docs"),
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
        self.docs_tables_wrapper.delete_row(row_index)


    # # 预览上传文件（调用worker）
    # @set_error_wrapper
    # def preview_table_button(self, row_index, table_type, *args, **kwargs):
    #     # 读取文件进行上传
    #     df_config = self.get_df_config_by_row_index(row_index, table_type)
    #     df_config["nrows"] = 10  # 实际读取的行数
    #     params = {
    #         "stage": "preview_df",  # 第一阶段
    #         "df_config": df_config,  # 上传的所有文件名
    #     }
    #     self.worker.add_params(params).start()
    #     self.tip_loading.set_titles(["预览表格.", "预览表格..", "预览表格..."]).show()
    #
    # @set_error_wrapper
    # def custom_preview_df(self, preview_result):
    #     df = preview_result.get("df")
    #     status_msg = preview_result.get("status_msg")
    #     max_rows_to_show = 10
    #     if len(df) >= max_rows_to_show:
    #         extra = [f'...省略剩余行' for _ in range(df.shape[1])]
    #         new_row = pd.Series(extra, index=df.columns)
    #         # 截取前 max_rows_to_show 行，再拼接省略行信息
    #         df = pd.concat([df.head(max_rows_to_show), pd.DataFrame([new_row])], ignore_index=True)
    #     self.tip_loading.hide()
    #     self.set_status_text(status_msg)
    #     self.table_modal(df, size=(400, 200))
    #
    # @set_error_wrapper
    # def get_df_by_row_index(self, row_index, table_type, nrows=None, *args, **kwargs):
    #     df_config = self.get_df_config_by_row_index(row_index, table_type)
    #     df_config["nrows"] = nrows
    #     return read_excel_file_with_multiprocessing([df_config])[0]
    #
    # @set_error_wrapper
    # def get_df_config_by_row_index(self, row_index, table_type, *args, **kwargs):
    #     if table_type == "main":
    #         table_wrapper = self.main_tables_wrapper
    #     else:
    #         table_wrapper = self.help_tables_wrapper
    #     path = table_wrapper.get_cell_value(row_index, 4)
    #     sheet_name = table_wrapper.get_cell_value(row_index, 1)  # 工作表
    #     row_num_for_column = table_wrapper.get_cell_value(row_index, 2)  # 列所在行
    #     return {
    #         "path": path,
    #         "sheet_name": sheet_name,
    #         "row_num_for_column": row_num_for_column,
    #     }
    #
    @set_error_wrapper
    def add_action(self, *args, **kwargs):
        """
        顺序 ｜ 类型 ｜ 动作 ｜ 动作内容 ｜ 操作按钮
        :return:
        """
        action_type = self.add_action_combo.currentText()  #  定位 | 选择 | 修改 | 合并
        action_type_obj = ACTION_MAPPING.get(action_type, {})

        self.actions_table_wrapper.add_rich_widget_row([
                {
                    "type": "readonly_text",
                    "value": action_type,  # 类型
                }, {
                    "type": "dropdown",
                    "values": [i.get("name") for i in action_type_obj.get("children")],  # 可选的动作
                }, {
                    "type": "editable_txt",  # 动作内容
                    "value": "",
                }, {
                    "type": "button_group",
                    "values": [
                        # {
                        #     "value": "向上移动",
                        #     "onclick": lambda row_index, col_index, row: self.actions_table_wrapper.swap_rows(
                        #         row_index, row_index+1),
                        # },
                        # {
                        #     "value": "向下移动",
                        #     "onclick": lambda row_index, col_index, row: self.actions_table_wrapper.swap_rows(
                        #         row_index, row_index-1),
                        # },
                        {
                            "value": "删除",
                            "onclick": lambda row_index, col_index, row: self.actions_table_wrapper.delete_row(
                                row_index),
                        },
                    ],
                }
            ])

    # @set_error_wrapper
    # def custom_after_add_condition(self, add_condition_result):
    #     status_msg = add_condition_result.get("status_msg")
    #     df_main_columns = add_condition_result.get("df_main_columns")
    #     table_name = add_condition_result.get("table_name")
    #     df_help_columns = add_condition_result.get("df_help_columns")
    #
    #     # 获取上一个条件的主表匹配列
    #     default_main_col_index = None
    #     if self.conditions_table_wrapper.row_length() > 0:
    #         default_main_col = self.conditions_table_wrapper.get_cell_value(
    #             self.conditions_table_wrapper.row_length() - 1, 0)
    #         if default_main_col in df_main_columns:
    #             default_main_col_index = df_main_columns.index(default_main_col)
    #
    #     # 构造级连选项
    #     # first_as_none = {"label": "***不从辅助表增加列***"}
    #     cascader_options = [{"label": NO_CATCH_COLS_OPTION}]
    #     for option in df_help_columns:
    #         column_option = {"label": option, "children": [
    #             {"label": ADD_COL_OPTION},
    #             {"label": MAKEUP_MAIN_COL, "children": [
    #                 {"label": main_label} for main_label in df_main_columns
    #             ]}
    #         ]}
    #         cascader_options.append(column_option)
    #     self.conditions_table_wrapper.add_rich_widget_row([
    #         {
    #             "type": "dropdown",
    #             "values": df_main_columns,  # 主表匹配列
    #             "cur_index": default_main_col_index if default_main_col_index is not None else 0,
    #         }, {
    #             "type": "readonly_text",
    #             "value": table_name,  # 辅助表
    #         }, {
    #             "type": "dropdown",
    #             "values": df_help_columns,  # 辅助表匹配列
    #         }, {
    #             "type": "dropdown",
    #             "values": [IGNORE_NOTHING, IGNORE_PUNC, IGNORE_CHINESE_PAREN, IGNORE_ENGLISH_PAREN],  # 重复值策略
    #             "cur_index": 1,  # 默认只忽略所有中英文标点符号
    #             "options": {
    #                 "multi": True,
    #                 "bg_colors": [COLOR_YELLOW] + [None] * 4,
    #                 "first_as_none": True,
    #             }
    #             # }, {
    #             #     "type": "dropdown",
    #             #     "values": ["***不从辅助表增加列***", *df_help_columns],  # 列：从辅助表增加
    #             #     "options": {
    #             #         "multi": True,
    #             #         "bg_colors": [COLOR_YELLOW] + [None] * len(df_help_columns),
    #             #         "first_as_none": True,
    #             #     }
    #
    #         }, {
    #             "type": "dropdown",
    #             "values": cascader_options,  # 列：从辅助表增加
    #             "cur_index": [0],
    #             "options": {
    #                 "cascader": True,
    #                 # "bg_colors": [COLOR_YELLOW] + [None] * len(df_help_columns),
    #                 "first_as_none": True,
    #             }
    #
    #         }, {
    #             "type": "editable_text",  # 列：匹配情况
    #             "value": " ｜ ".join(MATCH_OPTIONS),
    #         }, {
    #             "type": "readonly_text",  # 列：系统匹配到的行数
    #             "value": "匹配到的行数",
    #         }, {
    #             "type": "dropdown",
    #             "values": ["***不删除行***", *MATCH_OPTIONS],
    #             "options": {
    #                 "multi": True,
    #                 "first_as_none": True,
    #             }
    #         }, {
    #             "type": "button_group",
    #             "values": [
    #                 {
    #                     "value": "删除",
    #                     "onclick": lambda row_index, col_index, row: self.conditions_table_wrapper.delete_row(
    #                         row_index),
    #                 },
    #             ],
    #
    #         }
    #     ])
    #     self.tip_loading.hide()
    #     self.set_status_text(status_msg)
    #
    # @set_error_wrapper
    # def check_table_condition(self, row_index, row, *args, **kwargs):
    #     df_main = self.get_df_by_row_index(0, "main")
    #     df_help = self.get_df_by_row_index(row_index, "help")
    #     main_col = row["主表匹配列"]
    #     help_col = row["辅助表匹配列"]
    #
    #     duplicate_info = check_match_table(df_main, [{
    #         "df": df_help,
    #         "match_cols": [{
    #             "main_col": main_col,
    #             "match_col": help_col,
    #         }],
    #     }])
    #     if duplicate_info:
    #         dup_values = ", ".join([str(i.get("duplicate_cols", {}).get("cell_values", [])[0]) for i in duplicate_info])
    #         msg = "列：{}\t重复值{}".format(help_col, dup_values)
    #         self.modal("warn", f"经过检查辅助表存在重复: \n{msg}")
    #     else:
    #         self.modal("info", "检查通过，辅助表不存在重复")
    #
    @set_error_wrapper
    def run(self, *args, **kwargs):
        df_docs = self.docs_tables_wrapper.get_data_as_df()
        if len(df_docs) == 0:
            return self.modal(level="warn", msg="请先上传文档")
        df_actions = self.actions_table_wrapper.get_data_as_df()
        if len(df_actions) == 0:
            return self.modal(level="warn", msg="请先添加动作流")

        params = {
            "stage": "run",  # run
            "df_docs": df_docs,  # 文档的路径
            "df_actions": df_actions,  # 动作流
            # "result_table_wrapper": self.result_table_wrapper,  # 结果表的wrapper
        }

        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["文档处理.", "文档处理..", "文档处理..."]).show()

    @set_error_wrapper
    def custom_after_run(self, run_result):
        tip = run_result.get("tip")
        status_msg = run_result.get("status_msg")
        duration = run_result.get("duration")
        # self.detail_match_info = run_result.get("detail_match_info")
        # self.overall_match_info = run_result.get("overall_match_info")
        # self.matched_df = run_result.get("matched_df")
        # self.odd_cols_index = run_result.get("odd_cols_index")
        # self.even_cols_index = run_result.get("even_cols_index")
        # self.overall_cols_index = run_result.get("overall_cols_index")
        # self.match_for_main_col = run_result.get("match_for_main_col")

        self.result_detail_text.setText(tip)
        self.tip_loading.hide()
        self.set_status_text(status_msg)
        return self.modal(level="info", msg=f"✅文档处理成功，共耗时：{duration}秒")

    @set_error_wrapper
    def update_preview_col_num(self, step):
        new_num = int(self.preview_col_num_text.text()) + step
        if 0 < new_num < 6:
            self.preview_col_num_text.setText(str(new_num))



    # @set_error_wrapper
    # def show_result_detail_info(self, *args, **kwargs):
    #     if not self.detail_match_info:
    #         return self.modal(level="warn", msg="请先执行")
    #     msg_list = []
    #     data = []
    #     for k, v in self.detail_match_info.items():
    #         duration = round(v.get("time_cost") * 1000, 2)
    #         match_percent = len(v.get('match_index_list')) / (
    #                     len(v.get('match_index_list')) + len(v.get('unmatch_index_list')))
    #         unmatch_percent = len(v.get('unmatch_index_list')) / (
    #                     len(v.get('match_index_list')) + len(v.get('unmatch_index_list')))
    #         delete_percent = len(v.get('delete_index_list')) / (
    #                     len(v.get('match_index_list')) + len(v.get('unmatch_index_list')))
    #         data.append({
    #             "表名": k,
    #             # "耗时": f"{duration}s",
    #             "匹配行数": f"{len(v.get('match_index_list'))}（{round(match_percent * 100, 2)}%）",
    #             "未匹配行数": f"{len(v.get('unmatch_index_list'))}（{round(unmatch_percent * 100, 2)}%）",
    #             "需要删除行数": f"{len(v.get('delete_index_list'))}（{round(delete_percent * 100, 2)}%）",
    #         })
    #     self.table_modal(pd.DataFrame(data), size=(500, 200))
    #
    # @set_error_wrapper
    # def view_result(self, *args, **kwargs):
    #     if not self.detail_match_info:
    #         return self.modal(level="warn", msg="请先执行")
    #
    #     table_widget_container = TableWidgetWrapper()
    #     params = {
    #         "stage": "view_result",  # 阶段：预览大表格
    #         "matched_df": self.matched_df,  # 匹配结果
    #         "table_widget_container": table_widget_container,  # 匹配结果
    #         "odd_cols_index": self.odd_cols_index,  # 偶数辅助表相关列的索引
    #         "even_cols_index": self.even_cols_index,  # 奇数辅助表相关列的索引
    #         "overall_cols_index": self.overall_cols_index,  # 综合列的索引（最后两列）
    #         "match_for_main_col": self.match_for_main_col,  # 综合列的索引（最后两列）
    #     }
    #     self.worker.add_params(params).start()
    #     self.tip_loading.set_titles(["生成预览结果.", "生成预览结果..", "生成预览结果..."]).show()
    #
    # @set_error_wrapper
    # def custom_view_result(self, view_result):
    #     table_widget_wrapper = view_result.get("table_widget_wrapper")
    #     status_msg = view_result.get("status_msg")
    #     self.tip_loading.hide()
    #     self.set_status_text(status_msg)
    #     self.table_modal(
    #         table_widget_wrapper, size=(1200, 1000)
    #     )
    #
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
    #
    # @set_error_wrapper
    # def custom_after_download(self, after_download_result):
    #     status_msg = after_download_result.get("status_msg")
    #     duration = after_download_result.get("duration")
    #     file_path = after_download_result.get("file_path")
    #     self.set_status_text(status_msg)
    #     self.tip_loading.hide()
    #     return self.modal(level="info", msg=f"✅下载成功，共耗时：{duration}秒", funcs=[
    #         # QMessageBox.ActionRole | QMessageBox.AcceptRole | QMessageBox.RejectRole
    #         # QMessageBox.DestructiveRole | QMessageBox.HelpRole | QMessageBox.YesRole | QMessageBox.NoRole
    #         # QMessageBox.ResetRole | QMessageBox.ApplyRole
    #
    #         {"text": "打开所在文件夹", "func": lambda: open_file_or_folder_in_browser(os.path.dirname(file_path)),
    #          "role": QMessageBox.ActionRole},
    #         {"text": "打开文件", "func": lambda: open_file_or_folder_in_browser(file_path),
    #          "role": QMessageBox.ActionRole},
    #     ])

    def custom_update_progress(self, value, *args, **kwargs):
        self.run_progress_bar.setValue(int(value * 100))  # 0-100的整数