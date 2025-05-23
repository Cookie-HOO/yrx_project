import functools
import json
import shutil
import traceback
import typing

import pandas as pd
from PyQt5.QtCore import pyqtSignal, QThread, QTimer, QTime, Qt
from PyQt5.QtWidgets import QWidget, QLabel, QMainWindow, QMessageBox, QListWidget, QFileDialog, QApplication, \
    QListView, QTextBrowser, QTableWidget, QDialog, QVBoxLayout, QTableWidgetItem, QHBoxLayout, QPushButton
from PyQt5.QtGui import QPixmap

from yrx_project.client.const import *
from yrx_project.client.utils.code_widget import CodeDialog
from yrx_project.client.utils.exception import ClientWorkerException
from yrx_project.client.utils.message_widget import TipWidgetWithCountDown, MyQMessageBox, TipWidgetWithLoading, \
    FormModal
from yrx_project.utils.file import get_file_name_without_extension, copy_file, get_file_name_with_extension, make_zip
from yrx_project.utils.logger import logger_sys_error
from yrx_project.utils.time_obj import TimeObj


def set_error_wrapper(func):
    @functools.wraps(func)
    def wrapper(self, *args, **kwargs):
        try:
            return func(self, *args, **kwargs)
        except Exception as e:
            self.statusBar.showMessage(f"❌执行报错: {e}")
            return None
    return wrapper


class Background(QWidget):
    def __init__(self, parent=None):
        super(Background, self).__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.setGeometry(100, 100, 500, 500)
        self.setWindowTitle('Background Demo')

        # 设置背景图片
        self.background = QLabel(self)
        self.background.setPixmap(QPixmap(STATIC_FILE_PATH.format(file="background1.jpg")))
        self.background.setGeometry(0, 0, 500, 500)


class BaseWorker(QThread):
    """以signal结尾的信号发送时，会自动调用 MainWindow中定义的方法，如
    xxx_signal.emit()，会自动调用xxx方法

    """
    # 生命周期信号
    after_start_signal = pyqtSignal()
    refresh_signal = pyqtSignal(str)  # refresh的文本,可以显示在状态栏中
    before_finished_signal = pyqtSignal()
    # 元素操作信号
    modal_signal = pyqtSignal(str, str)  # info|warn|error, msg  其中error会终止程序
    clear_element_signal = pyqtSignal(str)  # clear
    append_element_signal = pyqtSignal(str, str)  # add
    set_element_signal = pyqtSignal(str, str)  # clear + add、

    #
    hide_tip_loading_signal = pyqtSignal()


    # 链式添加参数
    def add_param(self, param_key, param_value):
        setattr(self, param_key, param_value)
        return self

    def add_params(self, param_dict):
        for param_key, param_value in param_dict.items():
            setattr(self, param_key, param_value)
        return self

    # 获取参数
    def get_param(self, key):
        return getattr(self, key)

    # run的wrapper: 开始和结束报错的生命周期
    def run(self):
        self.after_start_signal.emit()
        try:
            self.my_run()
        except ClientWorkerException as e:
            return self.modal_signal.emit("error", str(e))
        except Exception as e:
            self.refresh_signal.emit(f"❌执行失败：{str(e)}")
            return self.modal_signal.emit("error", traceback.format_exc())
        self.before_finished_signal.emit()

    @logger_sys_error
    def my_run(self):
        raise NotImplementedError


class BaseWindow(QMainWindow):
    ############  元素操作: 直接调用, 或者是worker发送事件的消费者 ############
    # 清除元素
    def clear_element(self, element):
        ele = getattr(self, element)
        if ele is None:
            return self.modal("warn", f"no such element: {element}")
        ele.clear()

    # 元素追加内容
    def append_element(self, element, item):
        pass

    # 元素覆盖内容
    def set_element(self, element, item):
        self.clear_element(element)
        ele = getattr(self, element)
        if ele is None:
            self.modal("warn", f"no such element: {element}")
            return
        if isinstance(ele, QListWidget):
            ele.addItems(json.loads(item))

    ############ 组件封装 ############
    def modal(self, level, msg, title=None, done=None, **kwargs):
        """
        :param level:
        :param msg:
        :param title:
        :param done:
        :param kwargs
            count_down
        :return:
        """
        title = title or level
        if level == "error" and done is None:
            done = True

        if level == "info":
            if done:
                self.done = True
            if kwargs.get("funcs") or kwargs.get("width") or kwargs.get("height"):
                MyQMessageBox(title=title, msg=msg, width=kwargs.get("width"), height=kwargs.get("height"), funcs=kwargs.get("funcs"))
            else:
                QMessageBox.information(self, title, msg)
        elif level == "warn":
            if done:
                self.done = True
            QMessageBox.warning(self, title, msg)
        elif level == "error":
            if done:
                self.set_status_failed()
                self.done = True
            QMessageBox.critical(self, title, msg)
        elif level == "check_yes":  # 只要yes（点击No或者关闭，reply都是一个值）
            default_yes_or_false = kwargs.get("default") or "yes"
            default = QMessageBox.No
            if default_yes_or_false.lower() == "yes":
                default = QMessageBox.Yes
            reply = QMessageBox.question(
                self, title, msg, QMessageBox.Yes | QMessageBox.No, default)
            return reply == QMessageBox.Yes
        elif level == "tip":
            count_down = kwargs.get("count_down", 3)
            TipWidgetWithCountDown(msg=msg, count_down=count_down)
        elif level == "loading":
            tip_with_loading = TipWidgetWithLoading()
            return tip_with_loading
        elif level == "form":
            fields_config = kwargs.get("fields_config", [])
            """
            fields_config = [
                {
                    "id": "checkbox",
                    "type": "checkbox",
                    "label": "检测到「序号」列，重置每个拆分结果的「序号」列",
                    "default": True,
                    "show_if": lambda : "序号" in raw_df.columns,
                },
                {
                    "id": "name",
                    "type": "editable_text",
                    "label": "请输入姓名：",
                    "default": "张三",
                    "placeholder": "请输入姓名",
                    "limit": lambda x: "不能为空" if len(x) == 0 else "",
                },
                {
                    "id": "tip",
                    "type": "tip",
                    "label": "一行提示文本，仅用作提示",
                },
            ]
            """
            # 一个表单的弹窗，将 fields_config 转化成一个可提交表单
            # 默认是yes，还有取消按钮
            reply, result = FormModal.show_form(title=title, msg=msg, fields_config=fields_config)
            """
            result: {
                "name": "张三",  # 用户输入的内容
                "tip": "一行提示文本，仅用作提示",
            }
            """
            return reply, result


    def table_modal(self, table_widget_or_wrapper_or_df, size=None):
        """
        弹出一个表格
        :param table_widget_or_wrapper_or_df:
        :param size: (1200, 1000)
        :return:
        """
        from yrx_project.client.utils.table_widget import TableWidgetWrapper
        table_widget = None
        if isinstance(table_widget_or_wrapper_or_df, QTableWidget):
            table_widget = table_widget_or_wrapper_or_df
        elif isinstance(table_widget_or_wrapper_or_df, TableWidgetWrapper):
            table_widget = table_widget_or_wrapper_or_df.table_widget
        elif isinstance(table_widget_or_wrapper_or_df, pd.DataFrame):
            table_widget = TableWidgetWrapper().fill_data_with_color(table_widget_or_wrapper_or_df).table_widget
        dialog = QDialog()
        dialog.setWindowTitle("表格预览")
        layout = QVBoxLayout(dialog)
        layout.addWidget(table_widget)
        dialog.setLayout(layout)
        if size:
            dialog.resize(*size)
        dialog.exec_()

    def code_modal(self, apply_func, language="python", init_code=None):
        dialog = CodeDialog(language=language, apply_func=apply_func, init_code=init_code)
        dialog.exec_()

    def list_modal(self, list_items, cur_index=0, msg=None,
                   confirm_button="确定", cancel_button="取消"):
        """
        弹出一个列表选择对话框
        :param list_items: 列表项数据（字符串列表）
        :param cur_index: 初始选中索引
        :param msg: 顶部提示信息
        :param confirm_button: 确认按钮文本
        :param cancel_button: 取消按钮文本
        :return: (selected_index, confirmed)
        """
        # 创建对话框
        dialog = QDialog()
        dialog.setWindowTitle("请选择")

        # 主布局
        layout = QVBoxLayout(dialog)

        # 添加提示信息
        if msg:
            lbl_msg = QLabel(msg)
            lbl_msg.setAlignment(Qt.AlignCenter)
            layout.addWidget(lbl_msg)

        # 创建列表控件
        list_widget = QListWidget()
        list_widget.addItems(list_items)

        # 设置初始选中项（防越界处理）
        if 0 <= cur_index < len(list_items):
            list_widget.setCurrentRow(cur_index)
        layout.addWidget(list_widget)

        # 创建按钮布局
        btn_layout = QHBoxLayout()

        # 确认按钮
        btn_confirm = QPushButton(confirm_button)
        btn_confirm.clicked.connect(dialog.accept)
        btn_layout.addWidget(btn_confirm)

        # 取消按钮
        btn_cancel = QPushButton(cancel_button)
        btn_cancel.clicked.connect(dialog.reject)
        btn_layout.addWidget(btn_cancel)

        layout.addLayout(btn_layout)

        # 处理对话框结果
        result = dialog.exec_()

        # 返回值处理
        if result == QDialog.Accepted:
            return list_widget.currentRow(), True
        else:
            return (cur_index if 0 <= cur_index < len(list_items) else -1), False

    # 上传
    def upload_file_modal(self, patterns=("Excel Files", "*.xlsx"), multi=False, required_base_name_list=None, copy_to: str = None) -> typing.Union[str, list, None]:
        """
        :param patterns:
            [(pattern_name, pattern), (pattern_name, pattern)]  ("Excel Files", "*.xls*")
        :param multi: 是否支持多选
        :param required_base_name_list: 上传的文件必須有的文件名,，注意这个参数的元素不含后缀且没有路径，只有文件名
        :param copy_to: 如果制定了这个参数，会将上传的文件一并拷贝到这个路径
        :return:
        """
        if copy_to:
            self.clear_tmp_and_copy_important(tmp_path=copy_to)
        if len(patterns) == 2 and isinstance(patterns[0], str):
            patterns = [patterns]
        options = QFileDialog.Options()
        pattern_str = ";;".join([f"{pattern_name} ({pattern})" for pattern_name, pattern in patterns])
        func = QFileDialog.getOpenFileNames if multi else QFileDialog.getOpenFileName
        file_name_or_list, _ = func(self, "QFileDialog.getOpenFileName()", "", pattern_str, options=options)
        if not file_name_or_list:
            return None
        # 处理必须要包含某些required_base_name的情况
        required_list = required_base_name_list or []
        file_name_list = file_name_or_list if isinstance(file_name_or_list, list) else [file_name_or_list]
        file_name_base_names = [get_file_name_without_extension(file_name) for file_name in file_name_list]
        for required in required_list:
            if required not in file_name_base_names:
                self.modal("warn", f"请包含{required}文件")
                return []

        # 如果指定了路径，将所有文件拷贝过去
        if copy_to:
            for file_name in file_name_list:
                new_path = os.path.join(copy_to, get_file_name_with_extension(file_name))
                copy_file(file_name, new_path)
        if isinstance(file_name_or_list, str):
            file_name_or_list = [file_name_or_list]
        return file_name_or_list

    # 下载
    def download_file_modal(self, default_name: str):
        options = QFileDialog.Options()
        suffix = default_name.split(".")[-1]
        file_path, _ = QFileDialog.getSaveFileName(self, "QFileDialog.getSaveFileName()", default_name,
                                                   f"All Files (*);;Text Files (*.{suffix})", options=options)
        return file_path

    def download_zip_from_path(self, path, default_topic):
        file_path = self.download_file_modal(f"{TimeObj().time_str}_{default_topic}.zip")
        if not file_path:
            return False, None
        make_zip(path, file_path.rstrip(".zip"))
        return True, file_path

    # copy_file(DAILY_REPORT_RESULT_TEMPLATE_PATH, filePath)

    # 复制
    @staticmethod
    def copy2clipboard(text: str):
        QApplication.clipboard().setText(text)

    @staticmethod
    def clear_tmp_and_copy_important(tmp_path=None, important_path=None):
        # 1. 创建路径
        shutil.rmtree(tmp_path, ignore_errors=True)
        os.makedirs(tmp_path, exist_ok=True)
        if important_path:
            # 2. 拷贝关键文件到tmp路径
            for file in os.listdir(important_path):
                if not file.startswith("~") and (file.endswith("xlsx") or file.endswith("xlsm")):
                    old_path = os.path.join(important_path, file)
                    new_path = os.path.join(tmp_path, file)
                    copy_file(old_path, new_path)

    ############ wrapper类函数: 函数式编程思想,减少代码 ############
    def func_modal_wrapper(self, msg, func, *args, **kwargs):
        func(*args, **kwargs)
        self.modal("info", msg)

    def modal_func_wrapper(self, limit, warn_msg, func, *args, **kwargs):
        if not limit:
            return self.modal("warn", warn_msg)
        func(*args, **kwargs)


class WindowWithMainWorkerBarely(BaseWindow):
    """主窗口的一种模式
    在这个窗口中,存在一个主任务,窗口的设计都围绕这个主任务进行
    """
    SIGNAL_SUFFIX = "_signal"

    def __init__(self):
        super(WindowWithMainWorkerBarely, self).__init__()
        # 将worker的signal自动注册上handler
        self.tip_loading = self.modal(level="loading", titile="加载中...", msg=None)

        self.worker = self.register_worker()
        for worker_signal in self.worker.__dir__():
            if worker_signal.endswith(self.SIGNAL_SUFFIX):
                worker_handler = worker_signal[:-len(self.SIGNAL_SUFFIX)]
                worker_signal = getattr(self.worker, worker_signal)
                worker_signal.connect(getattr(self, worker_handler))
        # 计时器
        self.timer = QTimer()
        self.timer.timeout.connect(self.refresh_during_worker)
        self.start_time = None
        # 任务状态与显示
        self.status_bar_text = None  # worker中发出来后,绑定到这个变量,被statusbar更新
        self.__status = None

    # 注册一个Worker
    def register_worker(self) -> BaseWorker:
        pass

    ############  生命周期状态 ############
    @property
    def is_empty_status(self):
        return self.__status is None  # 原始状态

    @property
    def is_init(self):
        return self.__status == "init"

    @property
    def is_running(self):
        return self.__status == "running"

    @property
    def is_done(self):
        return self.__status in ["success", "failed"]

    @property
    def is_success(self):
        return self.__status == "success"

    @property
    def is_failed(self):
        return self.__status == "failed"

    def set_status_empty(self):
        self.timer.stop()
        self.__status = None

    def set_status_init(self):
        self.__status = "init"

    def set_status_running(self):
        self.__status = "running"

    def set_status_success(self):
        elapsed_time = self.start_time.secsTo(QTime.currentTime())
        # self.statusBar.showMessage(f"Success: Last for: {elapsed_time} seconds")
        self.timer.stop()
        self.__status = "success"
        # self.modal("info", title="Finished", msg=f"执行完成,共用时{self.start_time.secsTo(QTime.currentTime())}秒")

    def set_status_failed(self):
        self.timer.stop()
        self.__status = "failed"
        # if self.start_time:
        #     elapsed_time = self.start_time.secsTo(QTime.currentTime())
        #     self.statusBar.showMessage(f"Failed: Last for: {elapsed_time} seconds")


    ############  生命周期: 直接调用, 或者是worker发送事件的消费者 ############
    # 生命周期: worker任务启动
    def after_start(self):
        self.set_status_running()
        self.start_time = QTime.currentTime()
        self.timer.start(1)  # 更新频率为1毫秒

    # 生命周期: worker停止前每秒刷新的内容
    def refresh_during_worker(self):  # 计时器停止,自然就没人调用了,不用判断状态
        # elapsed_time = self.start_time.secsTo(QTime.currentTime())
        # self.status_text = f"Last for: {elapsed_time} seconds :: {self.refresh_text}"

        self.set_status_text(self.status_bar_text)  # 设置状态文字

    def refresh(self, status_bar_text):
        self.status_bar_text = status_bar_text

    # 生命周期: worker停止前
    def before_finished(self):
        self.set_status_success()

    def set_status_text(self, text):
        if isinstance(text, str):
            self.statusBar.showMessage(text)

    def hide_tip_loading(self):
        self.tip_loading.hide()
