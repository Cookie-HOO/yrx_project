import os

from yrx_project.scene.process_docs.base import ActionContext, CommandManager
from yrx_project.scene.process_docs.action_types import action_types
from yrx_project.utils.file import get_file_name_with_extension, copy_file


class ActionParser:

    def __init__(self):
        self.steps = []

    @staticmethod
    def parse(actions) -> CommandManager:
        """
        [
            {"action_id": "search_first_after", "action_params": {"content": "123"}},
            {"action_id": "move_left", "action_params": {"content": "123"}},
            {"action_id": "move_down", "action_params": {"content": "123"}},
            {"action_id": "select_current_cell", "action_params": {"content": "123"}},
            {"action_id": "replace_text", "action_params": {"content": "123"}},
            {"action_id": "merge_docs", "action_params": {}},
        ]
        """
        cm = CommandManager()
        for action_dict in actions:
            cm.add_command(action_types.init_command(action_dict.get("action_id"), action_dict.get("action_content")))
        return cm


class ActionProcessor:
    def __init__(self, config, after_each_action_func=None):
        self.command_manager = ActionParser.parse(config)
        self.command_containers = self.command_manager.command_containers
        self.context = ActionContext()
        self.context.command_manager = self.command_manager

        self.after_each_action_func = after_each_action_func

    def _process_command(self):
        ctx = self.context
        command_containers = self.command_containers[:]  # 复制一份，避免修改原始列表
        while command_containers:
            command_container = command_containers.pop(0)
            ctx.command_container = command_container

            if command_container.is_batch():
                # 1. batch 类型的任务：输入路径对齐到 自己的输出路径
                input_paths = [str(os.path.join(command_container.output_folder, get_file_name_with_extension(i))) for i in ctx.input_paths]
                for file_path, input_path in zip(ctx.input_paths, input_paths):
                    if not os.path.exists(file_path):
                        raise FileNotFoundError(f"File not found: {file_path}")
                    copy_file(file_path, input_path)
                ctx.input_paths = input_paths
                ctx.total_task_num = len(ctx.input_paths) * ctx.command_container.commands_num

                # 2. 执行任务：遍历文件，每个文件都需要执行batch下的所有命令
                for file_path in ctx.input_paths:  # 这里如何并发执行
                    ctx.file_path = file_path
                    ctx.into_file(file_path)
                    for command in command_container.commands:
                        ctx.command = command
                        command.run(ctx)
                        ctx.done_task()
                        if self.after_each_action_func is not None:
                            self.after_each_action_func(ctx)

            # 混合型任务: 不会修改inputs路径的文件，命令内部完成将ctx的input_path指向新的路径
            else:
                for command in command_container.commands:
                    ctx.total_task_num = 1
                    ctx.file_path = None
                    ctx.command = command
                    command.run(ctx)
                    ctx.done_task()
                    if self.after_each_action_func is not None:
                        self.after_each_action_func(ctx)

    def process(self, file_paths, **kwargs):
        self.context.init(file_paths)
        try:
            self._process_command()
        except Exception as e:
            print(f"Processing error: {str(e)}")
            raise
        finally:
            self.context.cleanup()
            log = self.context.log_df
        return

