import os
from typing import List, Dict, Optional, Callable, Generator


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
    def __init__(
        self,
        config: List[Dict],
        after_each_action_func: Optional[Callable[[Dict], None]] = None,
        debug_mode: bool = False
    ):
        self.command_manager = self._parse_config(config)
        self.command_containers = self.command_manager.command_containers
        self.context = ActionContext()
        self.after_each_action_func = after_each_action_func
        self.debug_mode = debug_mode
        self.execution_gen: Optional[Generator] = None
        self.current_container_idx = 0
        self.current_file_idx = 0
        self.current_cmd_idx = 0

    @staticmethod
    def _parse_config(config: List[Dict]):
        """解析动作配置（示例实现需补充）"""
        return ActionParser.parse(config)

    def init_context(self, input_paths: List[str]):
        """初始化执行上下文"""
        self.context.init(input_paths, command_manager=self.command_manager, debug_mode=True)

    def _batch_execution_generator(self) -> Generator:
        """批处理任务执行生成器"""
        ctx = self.context
        containers = self.command_containers[self.current_container_idx:]

        for cont_idx, container in enumerate(containers, start=self.current_container_idx):
            ctx.command_container = container
            self.current_container_idx = cont_idx

            if container.is_batch():
                # 初始化批处理环境
                input_paths = [
                    os.path.join(container.output_folder, os.path.basename(p))
                    for p in ctx.input_paths
                ]
                for src, dst in zip(ctx.input_paths, input_paths):
                    if not os.path.exists(src):
                        raise FileNotFoundError(f"Missing input: {src}")
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    copy_file(src, dst)
                ctx.input_paths = input_paths

                # 遍历文件
                for file_idx, file_path in enumerate(ctx.input_paths[self.current_file_idx:],
                                                   start=self.current_file_idx):
                    ctx.file_path = file_path
                    ctx.into_file(file_path)
                    self.current_file_idx = file_idx

                    # 遍历命令
                    cmds = container.commands[self.current_cmd_idx:]
                    for cmd_idx, cmd in enumerate(cmds, start=self.current_cmd_idx):
                        ctx.command = cmd
                        cmd.run(ctx)
                        ctx.done_task()
                        self.current_cmd_idx = cmd_idx
                        if self.after_each_action_func:
                            self.after_each_action_func(ctx.get_state())
                        yield  # 暂停点

                    self.current_cmd_idx = 0  # 重置命令索引
                self.current_file_idx = 0  # 重置文件索引

            else:
                # 普通任务执行
                cmds = container.commands[self.current_cmd_idx:]
                for cmd_idx, cmd in enumerate(cmds, start=self.current_cmd_idx):
                    ctx.command = cmd
                    cmd.run(ctx)
                    ctx.done_task()
                    self.current_cmd_idx = cmd_idx
                    if self.after_each_action_func:
                        self.after_each_action_func(ctx.get_state())
                    yield  # 暂停点
                self.current_cmd_idx = 0

    def process(self, input_paths: List[str]):
        """执行入口（非调试模式直接运行）"""
        self.init_context(input_paths)
        if not self.debug_mode:
            for _ in self._batch_execution_generator():
                pass

    def process_next(self) -> bool:
        """执行下一步，返回是否完成"""
        if not self.execution_gen:
            self.execution_gen = self._batch_execution_generator()
        try:
            next(self.execution_gen)
            return True  # 还有后续动作
        except StopIteration:
            self._cleanup()
            return False  # 全部完成

    def _cleanup(self):
        """清理资源"""
        self.context.cleanup()
        self.execution_gen = None
        self.current_container_idx = 0
        self.current_file_idx = 0
        self.current_cmd_idx = 0
