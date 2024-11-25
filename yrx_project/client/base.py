import functools
import json
import shutil
import traceback
import typing

import pandas as pd
from PyQt5.QtCore import pyqtSignal, QThread, QTimer, QTime
from PyQt5.QtWidgets import QWidget, QLabel, QMainWindow, QMessageBox, QListWidget, QFileDialog, QApplication, \
    QListView, QTextBrowser, QTableWidget, QDialog, QVBoxLayout, QTableWidgetItem
from PyQt5.QtGui import QPixmap

from yrx_project.client.const import *
from yrx_project.client.utils.exception import ClientWorkerException
from yrx_project.client.utils.message_widget import TipWidgetWithCountDown, MyQMessageBox, TipWidgetWithLoading
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
            default_yes_or_false = kwargs.get("default")
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
            return
        make_zip(path, file_path.rstrip(".zip"))

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
