import os
import time
import typing

from PyQt5 import uic
from PyQt5.QtCore import pyqtSignal

from yrx_project.client.base import WindowWithMainWorkerBarely, BaseWorker, set_error_wrapper
from yrx_project.client.const import UI_PATH
from yrx_project.client.utils.button_menu_widget import ButtonMenuWrapper
from yrx_project.client.utils.line_splitter import LineSplitterWrapper
from yrx_project.client.utils.table_widget import TableWidgetWrapper
from yrx_project.client.utils.tree_file_widget import TreeFileWrapper
from yrx_project.const import PROJECT_PATH, TEMP_PATH
from yrx_project.scene.process_docs.base import ActionContext
from yrx_project.scene.process_docs.const import SCENE_TEMP_PATH
from yrx_project.client.scene.docs_process_adapter import build_action_types_menu, \
    cleanup_scene_folder, has_content_in_scene_folder, build_action_suit_menu, ActionRunner
from yrx_project.utils.file import get_file_name_without_extension, get_file_detail, FileDetail, open_file_or_folder, \
    get_file_name_with_extension, copy_file
from yrx_project.utils.iter_util import find_repeat_items
from yrx_project.utils.time_obj import TimeObj


class Worker(BaseWorker):
    custom_after_upload_signal = pyqtSignal(dict)  # 自定义信号
    # custom_after_add_condition_signal = pyqtSignal(dict)  # 自定义信号
    custom_after_run_signal = pyqtSignal(dict)  # 自定义信号
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
            file_paths = self.get_param("file_paths")
            # 校验是否有同名文件
            file_details = [get_file_detail(file_path) for file_path in file_paths]
            all_base_name_list = [i.name_without_extension for i in file_details] + table_wrapper.get_data_as_df()["文档名称"].to_list()
            repeat_items = find_repeat_items(all_base_name_list)
            if repeat_items:
                repeat_items_str = '\n'.join(repeat_items)
                self.hide_tip_loading_signal.emit()
                return self.modal_signal.emit("warn", f"存在重复文件名，请修改后上传: \n{repeat_items_str}")

            check_same_name = time.time()
            # pages = get_docx_pages_with_multiprocessing(file_names)
            # read_file_time = time.time()
            status_msg = \
                f"✅上传{len(file_paths)}个文档成功，共耗时：{round(time.time() - start_upload_time, 2)}s："\
                f"校验文件名：{round(check_same_name - start_upload_time, 2)}s；"\
                # f"读取文件：{round(read_file_time - check_same_name, 2)}s；"\

            self.custom_after_upload_signal.emit({
                # "pages": pages,
                "file_details": file_details,
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

            # df_docs = self.get_param("df_docs")
            # df_actions = self.get_param("df_actions")
            action_runner: ActionRunner = self.get_param("action_runner")
            action_runner.after_each_action_func = lambda ctx: self.refresh_signal.emit(f"文档处理中: {ctx.get_show_msg()}")
            action_runner.run_actions()

            # 设置执行信息
            duration = round((time.time() - start_run_time), 2)
            tip = f"✅处理成功，共耗时：{duration}秒"

            status_msg = \
                f"✅处理成功，共耗时：{duration}秒"

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
                step1_help_info_button
                add_docs_button：添加word文档
                docs_table
                    文档名称 | 文件大小 | 修改时间 | 操作按钮 | __文档路径
            第二步：定义动作流
                step2_help_info_button
                action_tools_button：动作流导入导出
                add_action_button：设置匹配条件
                actions_table：动作流表格
                    类型 ｜ 动作 ｜ 动作内容 | 操作按钮 | __动作id
            第三步：执行
                step3_help_info_button
                run_button：执行按钮
                run_log_button：执行按钮
                result_detail_text：执行详情
                     🚫执行耗时：--毫秒
                download_result_button: 下载结果按钮
                result_tree：结果文件的树状结构
            第四步（可选）：调试：单步执行
                step4_help_info_button
                debug_button
                debug_next_button
                actions_with_log_table

        """

    help_info_text = """<html>
    <head>
        <title>多文档操作场景</title>
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
        <h4>第一步：上传：通过拖拽或点击上传文档后得到列表</h4>
        <div class="table-container">
            <div class="table-wrapper1">
                <table>
                    <tr>
                        <th>文档名称</th>
                        <th>文件大小</th>
                        <th>修改时间</th>
                        <th>操作按钮</th>
                    </tr>
                    <tr>
                        <td>第1篇文档</td>
                        <td>12.4kb</td>
                        <td>2025-01-01 11:11:11</td>
                        <td>|删除|</td>
                    </tr>
                    <tr>
                        <td>第2篇文档</td>
                        <td>12.4kb</td>
                        <td>2025-01-01 11:11:11</td>
                        <td>|删除|</td>
                    </tr>
                    <tr>
                       <td>第3篇文档</td>
                        <td>12.4kb</td>
                        <td>2025-01-01 11:11:11</td>
                        <td>|删除|</td>
                    </tr>
                </table>
            </div>
            <div class="table-wrapper1">
            <h4>第二步：定义：动作流</h4>
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
                        <td>|⬆︎|⬇︎|❌|</td>
                    </tr>
                    <tr>
                        <td>2</td>
                        <td>定位</td>
                        <td>向左移动</td>
                        <td> 1 </td>
                        <td>|⬆︎|⬇︎|❌|</td>
                    </tr>
                    <tr>
                        <td>3</td>
                        <td>选择</td>
                        <td>选择当前单元格</td>
                        <td> --- </td>
                        <td>|⬆︎|⬇︎|❌|</td>
                    </tr>
                    <tr>
                        <td>4</td>
                        <td>修改</td>
                        <td>文字替换</td>
                        <td> abc </td>
                        <td>|⬆︎|⬇︎|❌|</td>
                    </tr>
                    <tr>
                        <td>5</td>
                        <td>总体</td>
                        <td>合并成一个文档</td>
                        <td> -- </td>
                        <td>|⬆︎|⬇︎|❌|</td>
                    </tr>
                </table>
                <h4>第三步：执行</h4>
                <p>左下角展示执行后后的文件树：会按照文件夹进行组织，如</p>
                <p>1-batch：存放批量操作的内容，系统会自动合并所有批量操作（目前是除了混合文档之外的操作），批量操作结束后会保存文件</p>
                <p>2-mixing：存放混合文档后的结果，目前仅支持文档合并</p>
                <p>*前面的数字表示执行的顺序</p>
                <p>假设动作流设置如下</p>
                <p>1. 搜索并选中</p>
                <p>2. 修改字体</p>
                <p>3. 修改字号</p>
                <p>4. 合并所有文档</p>
                <p>那么最终会将第1-3步的结果文件生成到1-batch文件夹中，第4步的结果放到2-mixing中</p>
            </div>
        </div>
        <h4>结果：</h4>
    </body>
    </html>"""
    release_info_text = """
    v1.0.6: 实现基础版本的文档聚合
    - 实现多文档操作场景
        - 上传多个文档
        - 定义操作流
        - 批量执行
        - 单步调试等
    """

    # 第一步：上传文件的帮助信息
    step1_help_info_text = """
    第一步：上传文件
    1. 可点击按钮或拖拽文档到表格中：目前只支持docx格式
    2. 上传后展示文件名、大小、修改时间等
    """
    # 第二步：添加动作流的帮助信息
    step2_help_info_text = """
    第二步：添加工作流
    1. 点击添加，会显示添加的动作类型，目前支持：定位光标、光标位置插入、选择内容、修改选中内容、混合文档，五类操作
    2. 指定动作类型后，在动作中选择一个对应的动作
    3. 输入动作内容
    4. 操作按钮中可以：向上移动、向下移动、删除
    5. 工具按钮可以
        - 复制当前动作流到剪贴板，下次可以将这个内容导入
        - 导入剪贴板中的内容：如果剪贴板中存在合法的曾经导出的内容，可以再进行导入
        - 加载内置的预设动作流，目前系统预设了一套动作流，可以加载，后续有需要可以继续内置到系统中
    """
    # 第三步：执行与下载的帮助信息
    step3_help_info_text = """
    第三步：批量执行
    1. 点击执行后，按照指定的动作流进行执行
    2. 执行过程中或完成后，可以随时点击log按钮进行日志查看
    3. 进度条会按照大的阶段显示进度
        1-batch下的所有进度
    4. 结果文件树中
        可以右键点击某个文件：打开或者下载
        可以就见点击某个文件夹进行打开
    5. 可以下载完整的结果
    """
    step4_help_info_text = """
    第四步（可选）：调试
    1. 点击调试按钮后，会进行提示，确认后进入调试模式（会同步打开word，单步进行执行）
    2. 默认以文件树的第一个文件进行调试
        如果需要指定文件，右键按钮指定需要调试的文件
    3. 调试开始后，需要点击下一步才会执行
    4. 每次执行会在下方的表格展示当前执行的步骤
        执行完成的，绿色背景
        执行有警告的，黄色背景
        执行失败的，红色背景
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

        # 布局修改
        ## 1. 上下布局可移动
        # 在代码中设置 Splitter 样式
        self.splitter_design = LineSplitterWrapper(self.splitter)
        # 1. 批量上传文档
        # 1.1 按钮
        self.add_docs_button.clicked.connect(self.add_docs)
        # self.reset_button.clicked.connect(self.reset_all)
        # 1.2 表格
        self.docs_tables_wrapper = TableWidgetWrapper(
            self.docs_table, drag_func=self.docs_drag_drop_event).set_col_width(2, 150)  # 上传docs之后展示所有table的表格
        # self.help_tables_wrapper = TableWidgetWrapper(self.help_tables_table,
        #                                               drag_func=self.help_drag_drop_event)  # 上传table之后展示所有table的表格
        #
        # # 2. 添加动作流
        self.actions_table_wrapper = TableWidgetWrapper(self.actions_table).set_col_width(1, 320).set_col_width(3, 140)
        self.add_action_button_menu = ButtonMenuWrapper(
            self, self.add_action_button, build_action_types_menu(self.actions_table_wrapper)
        )
        self.action_suit_button_menu = ButtonMenuWrapper(
            self, self.action_tools_button, build_action_suit_menu(self.actions_table_wrapper)
        )

        # self.add_action_button.clicked.connect(self.add_action)
        #
        # # 3. 执行与下载
        # self.matched_df, self.overall_match_info, self.detail_match_info = None, None, None  # 用来获取结果
        # self.odd_cols_index, self.even_cols_index, self.overall_cols_index = None, None, None  # 用来标记颜色
        # self.match_for_main_col = None  # 主表匹配列的映射
        self.run_button.clicked.connect(self.run)
        self.tree_file_wrapper = TreeFileWrapper(
            self.result_tree, SCENE_TEMP_PATH,
            on_double_click=lambda f: open_file_or_folder(f),
            right_click_menu=[
                {"type": "menu_action", "name": "打开",
                 "func": lambda f: open_file_or_folder(f)},
                {"type": "menu_action", "name": "保存",
                 "func": self.right_click_menu_save_file},
            ],
            open_on_default=[]
        )


        # self.worker.custom_after_upload_signal.connect(self.custom_after_upload)
        # self.result_table_wrapper = TableWidgetWrapper(self.result_table)
        # self.result_detail_info_button.clicked.connect(self.show_result_detail_info)
        # # self.preview_result_button.clicked.connect(self.preview_result)
        self.download_result_button.clicked.connect(self.download_result)
        # self.view_result_button.clicked.connect(self.view_result)

        # 第四步骤：调试
        self.debug_current_step = None
        self.action_runner: typing.Optional[ActionRunner] = None
        self.debug_file_paths = []  # 用于debug的输入路径
        self.debug_button.clicked.connect(self.debug_run)
        self.debug_next_button.clicked.connect(self.debug_next)
        self.actions_with_log_table_wrapper = TableWidgetWrapper(self.actions_with_log_table, disable_edit=True).set_col_width(1, 320).set_col_width(3, 140)

    def right_click_menu_save_file(self, path):
        save_to = self.download_file_modal(TimeObj().time_str + get_file_name_with_extension(path))
        if save_to:
            copy_file(path, save_to)
            self.modal(level="info", msg="✅下载成功")

    def right_click_menu_preview_file(self, path):  #  TODO
        pass

    def register_worker(self):
        return Worker()

    def docs_drag_drop_event(self, file_paths):
        self.add_doc(file_paths)

    @set_error_wrapper
    def add_docs(self, *args, **kwargs):
        # 上传文件
        file_paths = self.upload_file_modal(["Word Files", "*.docx"], multi=True)
        if not file_paths:
            return
        self.add_doc(file_paths)
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
    def add_doc(self, file_paths):
        if isinstance(file_paths, str):
            file_paths = [file_paths]

        for file_path in file_paths:
            if not file_path.endswith(".docx"):
                return self.modal(level="warn", msg="仅支持docx文件")

        table_wrapper = self.docs_tables_wrapper

        # 读取文件进行上传
        params = {
            "stage": "upload",  # 第一阶段
            "file_paths": file_paths,  # 上传的所有路径名
            "table_wrapper": table_wrapper,
        }
        self.worker.add_params(params).start()
        self.tip_loading.set_titles(["上传文件.", "上传文件..", "上传文件..."]).show()

    # 上传文件的后处理
    @set_error_wrapper
    def custom_after_upload(self, upload_result):
        # pages = upload_result.get("pages")
        file_details: typing.List[FileDetail] = upload_result.get("file_details")
        status_msg = upload_result.get("status_msg")
        for file_detail in file_details:  # 辅助表可以一次传多个，主表目前只有一个
            self.docs_tables_wrapper.add_rich_widget_row([
                {
                    "type": "readonly_text",  # 文件名
                    "value": file_detail.name_without_extension,
                }, {
                    "type": "readonly_text",  # 文件大小
                    "value": str(file_detail.size_format),
                }, {
                    "type": "readonly_text",  # 修改时间
                    "value": str(file_detail.updated_at),
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
                    "value": file_detail.path,
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
    @set_error_wrapper
    def run(self, *args, **kwargs):
        df_docs = self.docs_tables_wrapper.get_data_as_df()
        if len(df_docs) == 0:
            return self.modal(level="warn", msg="请先上传文档")
        df_actions = self.actions_table_wrapper.get_data_as_df()
        if len(df_actions) == 0:
            return self.modal(level="warn", msg="请先添加动作流")

        if has_content_in_scene_folder():
            ok_or_not = self.modal(level="check_yes", msg=f"当前操作空间有上次执行的结果，是否继续（选择是，会清空之前的执行结果）", default="yes")
            if ok_or_not:
                cleanup_scene_folder()

        action_runner = ActionRunner(
            input_paths=df_docs["__文档路径"].to_list(),
            df_actions=df_actions,
        )

        params = {
            "stage": "run",  # run
            "action_runner": action_runner,
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
        self.tree_file_wrapper.force_refresh()
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
        if self.download_zip_from_path(path=SCENE_TEMP_PATH, default_topic="文档批处理"):
            self.modal(level="info", msg="✅下载成功")

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

    def debug_run(self):
        # 0. 至少上传了一个文件
        df_docs = self.docs_tables_wrapper.get_data_as_df()
        if len(df_docs) == 0:
            return self.modal(level="warn", msg="请先上传文档")
        df_actions = self.actions_table_wrapper.get_data_as_df()
        if len(df_actions) == 0:
            return self.modal(level="warn", msg="请先添加动作流")
        paths = df_docs["__文档路径"].to_list()
        base_names = [get_file_name_without_extension(i) for i in paths]

        # 1. 弹窗确认是否开始调试，以及列出所有文件让用户选择一个文件开始调试，默认第一个
        selected_index, yes_or_no = self.list_modal(
            list_items=base_names, cur_index=0, msg="指定输入文件开始调试（开始调试后无法修改动作流）"
        )
        if yes_or_no:
            self.debug_file_paths = base_names[selected_index]
        if len(self.debug_file_paths) == 0:
            return self.modal(level="warn", msg="没有输入文档，无法调试")

        # 2. 开始调试
        # 设置调试步骤（下一步的按钮需要）
        # 将第一步写到table中，增加一个行👉icon
        # 初始化执行环境：打开word
        self.actions_table_wrapper.disable_edit()
        action = df_actions.iloc[0,:]
        values = [
            action["类型"], action["动作"], action["动作内容"]
        ]
        self.actions_with_log_table_wrapper.add_rich_widget_row([
            {"type": "readonly_text", "value": str(i)} for i in values
        ])
        self.actions_with_log_table_wrapper.set_vertical_header(["👉"])
        action_runner = ActionRunner(
            input_paths=paths,
            df_actions=df_actions,
        )
        self.action_runner = action_runner
        try:
            action_runner.debug_actions()
            self.start_debug_mode()
        except Exception as e:
            self.actions_with_log_table_wrapper.clear()
            return self.modal(level="error", msg="初始化调试模式报错，尝试关闭相关文档再试"+str(e))

    def debug_next(self):
        if self.debug_current_step is None:
            return self.modal(level="warn", msg="请先进入调试模式")
        if self.action_runner is None:
            return self.modal(level="error", msg="未知错误，请重置")
        processor = self.action_runner.processor
        if processor is None:
            return self.modal(level="error", msg="未初始化processor")

        # 执行下一步
        self.action_runner.next_action_or_cleanup()
        self.debug_current_step += 1
        # 1. 修改这一步的日志
        log_row = processor.context.get_log_df().iloc[-1, :]  # 最后一行的日志
        level, msg = log_row["level"], log_row["msg"]
        self.actions_with_log_table_wrapper.set_cell_value(self.debug_current_step-1, 3, level+":" + msg)
        actions_log_df = self.actions_with_log_table_wrapper.get_data_as_df()
        result_list = [
            "✅" if row.startswith("info") else
            "⚠️" if row.startswith("warn") else
            "❌" if row.startswith("error") else None
            for row in actions_log_df["调试信息"]
        ]
        self.actions_with_log_table_wrapper.set_vertical_header(result_list)

        # 3. 增加下一步的记录
        # 检查是否已结束
        df_actions = self.actions_table_wrapper.get_data_as_df()
        if self.debug_current_step >= len(df_actions):
            return self.end_debug_mode()

        action = df_actions.iloc[self.debug_current_step,:]
        values = [
            action["类型"], action["动作"], action["动作内容"]
        ]
        self.actions_with_log_table_wrapper.add_rich_widget_row([
            {"type": "readonly_text", "value": str(i)} for i in values
        ])
        # 4. 设置表格样式
        # 设置行头，成功的 ✅，失败的❌，当前的 👉 不符合预期的 ⚠️
        self.actions_with_log_table_wrapper.set_vertical_header(result_list+["👉"])

    def start_debug_mode(self):
        self.debug_current_step = 0
        self.add_action_button_menu.disable_click(msg="当前处于调试模式，无法增加动作")
        self.action_suit_button_menu.disable_click([[0], [1]], msg="当前处于调试模式，无法增加动作")

    def end_debug_mode(self):
        self.modal(level="info", msg="✅调试结束")
        self.action_runner.next_action_or_cleanup()  # 如果执行完了，这一步就是清理，如果没有执行完，这一步就是向下执行
        self.add_action_button_menu.enable_click()
        self.action_suit_button_menu.enable_click()

