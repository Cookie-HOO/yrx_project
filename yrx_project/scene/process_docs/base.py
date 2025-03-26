import os
import shutil
import typing
from multiprocessing import Lock

from yrx_project.const import TEMP_PATH
from yrx_project.scene.process_docs.const import SCENE_TEMP_PATH

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
        self.doc = None
        self.selection = None

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

        # 4. 工具
        self.word = None  # 操作word的对象
        self.lock_task_num = Lock()
        self.lock_file_num = Lock()
        self.const = self.const_init()

    def const_init(self):
        from win32com.client import constants
        return {
            "SYMBOL_MAP": {
                "分页符": constants.wdPageBreak,
                "换行符": constants.wdLineBreak,
                "分节符": constants.wdSectionBreakNextPage,
                "制表符": constants.wdTab
            },
            "ALIGN_MAP": {
                "左对齐": constants.wdAlignParagraphLeft,
                "居中": constants.wdAlignParagraphCenter,
                "右对齐": constants.wdAlignParagraphRight,
                "两端对齐": constants.wdAlignParagraphJustify
            },
            "ROW_SPACING_MAP": {
                "倍数行距": constants.wdLineSpaceMultiple,
                "最小行距": constants.wdLineSpaceAtLeast,
                "固定行距": constants.wdLineSpaceExactly,
            },
            "SCOPE_MAP": {
                "行": constants.wdLine,
                "字符": constants.wdCharacter,
                "段落": constants.wdParagraph,
                "表格单元格": constants.wdCell,
                "页面": constants.wdPage,
                "整个文档": constants.wdStory
            },
            "BOUNDARY_CHECKS": {
                "cell_start": lambda s: s.Information(constants.wdWithInTable),
                "cell_end": lambda s: s.Information(constants.wdWithInTable),
                "page_start": lambda s: s.Information(constants.wdActiveEndPageNumber) > 0,
            },
            "BOUNDARY_ACTIONS" : {
                "line_start": (constants.wdLine, constants.wdMove),
                "line_end": (constants.wdLine, constants.wdExtend),
                "cell_start": (constants.wdCell, constants.wdMove),
                "cell_end": (constants.wdCell, constants.wdExtend),
                "page_start": (constants.wdPage, constants.wdMove),
                "page_end": (constants.wdPage, constants.wdExtend),
                "doc_start": (constants.wdStory, constants.wdMove),
                "doc_end": (constants.wdStory, constants.wdExtend)
            },
            "COLLAPSE_MAP": {
                "right": constants.wdCollapseEnd,
                "left": constants.wdCollapseStart,
            }
        }
    def done_task(self):  # 多进程级别保证同步
        with self.lock_task_num:
            self.done_task_num += 1

    def done_file(self):
        with self.lock_file_num:
            self.done_file_num += 1


class Command:
    action_type_id = None
    action_name = None

    def __init__(self, content, action_type_id, action_name, **kwargs):
        self.content = content
        self.action_type_id = action_type_id
        self.action_name = action_name

    def run(self, context: ActionContext):
        return self.office_word_run(context)

    def office_word_run(self, context: ActionContext):
        raise NotImplementedError

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



