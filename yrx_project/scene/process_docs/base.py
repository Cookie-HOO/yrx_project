import os
import shutil
import typing
from multiprocessing import Lock

import pandas as pd

from yrx_project.scene.process_docs.office_word_command_impl.command_impl_base import OfficeWordImplBase
from yrx_project.scene.process_docs.const import SCENE_TEMP_PATH
from yrx_project.scene.process_docs.office_word_command_impl.office_word_context import OfficeWordContext
from yrx_project.utils.file import get_file_name_without_extension

MIXING_TYPE_ID = "mixing"

ACTION_TYPE_MAPPING = {
    "locate": ("定位光标", "定位类操作，涉及搜索和光标移动"),
    "insert": ("光标位置插入", "插入类操作，如果跟在选择操作后面不起作用"),
    "select": ("选择内容", "选择类操作，选中一些内容，进行修改"),
    "update": ("修改选中内容", "修改类操作，修改选中的内容"),
    MIXING_TYPE_ID: ("混合文档", "混合类操作，输入n个文档，输出m个文档"),
}


class ActionContext:
    def __init__(self):
        # 1. 文件信息
        ## 输入文件路径
        self.init_input_paths: typing.List[str] = []  # 最最开始的输入路径（不变）
        self.input_paths: typing.List[str] = []  # 当前环境的输入文件路径（会随着阶段进行改变）

        ## 当前文件对象
        self.file_path: typing.Optional[str] = None  # 当前文件路径，如果是mixing类的任务，为None
        self.office_word_ctx = OfficeWordContext()

        # 2. 任务信息
        ## 所有任务
        self.command_manager: typing.Optional[CommandManager] = None
        ## 当前任务
        self.command_container: typing.Optional[CommandContainer] = None
        self.command: typing.Optional[Command] = None

        # 3. 日志信息
        self.done_task_num = 0
        self.total_task_num = 0
        self.done_file_num = 0
        self.columns = ["level", "msg", "file_name_without_extension", "command_ins"]
        self.log_df_ = pd.DataFrame(columns=self.columns)

        # 4. 工具
        self.lock_task_num = Lock()
        self.lock_file_num = Lock()
        self.lock_log = Lock()

    def get_show_msg(self):
        msg_list = [
            f"当前阶段: {self.command_container.step}/{len(self.command_manager.command_containers)}",
            f"阶段进度: {self.done_task_num / self.total_task_num * 100}%",
            # f"文件进度: {self.done_file_num}/{len(self.input_paths)}%",
            # f"当前动作: {self.command.action_name}; 当前文件: {self.file_name_without_extension}"
        ]
        return "; ".join(msg_list)

    def get_log_df(self):
        try:
            log_df = self.log_df_[["level", "msg", "file_name_without_extension"]]
            log_df["action_name"] = self.log_df_["command_ins"].apply(lambda x: x.action_name)
            log_df["action_content"] = self.log_df_["command_ins"].apply(lambda x: x.content)
            return log_df
        except AttributeError as e:
            raise AttributeError("log_df_ not initialized") from e

    def __getattr__(self, item):
        # 先检查是否是当前类应该处理的属性
        if item in ["word", "doc", "selection", "consts"]:  # 代理对这三个属性的访问：
            return getattr(self.office_word_ctx, item)
        raise AttributeError(f"'{self.__class__.__name__}' object has no attribute '{item}'")

    def init(self, file_paths, command_manager, debug_mode=False):
        # 初始化逻辑
        self.init_input_paths = file_paths
        self.input_paths = file_paths
        self.command_manager = command_manager
        self.command_manager.cleanup(file_paths)
        # 初始化 OfficeWordCtx
        self.office_word_ctx.init(debug_mode=debug_mode)

    def into_file(self, file_path):
        self.office_word_ctx.into_file(file_path)

    def quit_file(self):
        self.done_file()
        self.office_word_ctx.quit_file()

    def cleanup(self):
        self.office_word_ctx.cleanup()

    @property
    def file_name_without_extension(self):
        if self.file_path:
            return get_file_name_without_extension(self.file_path)
        return "---"

    def done_task(self):  # 多进程级别保证同步
        with self.lock_task_num:
            self.done_task_num += 1

    def done_file(self):
        with self.lock_file_num:
            self.done_file_num += 1

    def add_log(self, level, msg, file_name_without_extension, command_ins):
        new_row = pd.DataFrame(
            [[level, msg, file_name_without_extension, command_ins]],
            columns=self.columns,
        )
        with self.lock_log:
            self.log_df_ = pd.concat([self.log_df_, new_row], ignore_index=True)


class Command(OfficeWordImplBase):
    action_type_id = None
    action_name = None

    def __init__(self, action_type_id, action_type_name, action_name, action_id, action_content, **kwargs):
        super(Command, self).__init__()
        self.content = action_content
        self.action_type_id = action_type_id
        self.action_type_name = action_type_name
        self.action_id = action_id
        self.action_name = action_name

    def run(self, context: ActionContext):
        try:
            self.office_word_check(context)
            success, msg = self.office_word_run(context)
            if success:
                msg = "✅执行成功"
            context.add_log("info" if success else "warn", msg or "", context.file_name_without_extension, self)
            return success
        except Exception as e:
            context.add_log("error", str(e), context.file_name_without_extension, self)
            return False

    def is_mixing(self):
        return self.action_type_id == MIXING_TYPE_ID


class CommandContainer:
    def __init__(self):
        self.step = ""  # 在manager中，是排第几，从1开始
        self.commands = []

    @property
    def commands_num(self):
        return len(self.commands)

    def add_command(self, command: Command):
        self.commands.append(command)
        return self

    def set_step(self, step):
        self.step = step
        return self

    def is_batch(self):
        return isinstance(self, BatchCommandContainer)

    def is_mixing(self):
        return isinstance(self, MixingCommandContainer)

    @property
    def output_folder(self):
        return ""

    @property
    def step_and_name(self):
        return ""


class BatchCommandContainer(CommandContainer):
    @property
    def output_folder(self):
        return os.path.join(SCENE_TEMP_PATH, f"{self.step}-batch")

    @property
    def step_and_name(self):
        return f"{self.step}-批处理"


class MixingCommandContainer(CommandContainer):

    @property
    def output_folder(self):
        return os.path.join(SCENE_TEMP_PATH, f"{self.step}-mixing")

    @property
    def step_and_name(self):
        return f"{self.step}-混合处理"


class CommandManager:
    def __init__(self):
        self.command_containers: typing.List[CommandContainer] = []

    def add_command(self, command: Command):
        # 当前是mixing，一定需要添加新的container
        if command.is_mixing():
            self.command_containers.append(
                MixingCommandContainer().add_command(command).set_step(len(self.command_containers)+1)
            )
            return self

        # 当前是batch，可以考虑接着之前的
        if self.command_containers and self.command_containers[-1].is_batch():
            container = self.command_containers.pop()
        else:  # 新建一个container
            container = BatchCommandContainer().set_step(len(self.command_containers)+1)
        self.command_containers.append(container.add_command(command))
        return self

    def cleanup(self, file_paths):
        # 初始化准备执行所有命令
        ## 1. 清空各个container的output文件夹
        for container in self.command_containers:
            if os.path.exists(container.output_folder):
                shutil.rmtree(container.output_folder)
            os.makedirs(container.output_folder)

        # ## 2. 准备 0-init 文件夹
        # temp_file_paths = [os.path.join(TEMP_PATH, f"0-init", get_file_name_with_extension(i)) for i in file_paths]
        # for file_path, temp_file_path in zip(file_paths, temp_file_paths):
        #     if not os.path.exists(file_path):
        #         raise FileNotFoundError(f"File not found: {file_path}")
        #     copy_file(file_path, temp_file_path)



