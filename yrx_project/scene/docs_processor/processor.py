import os
import typing

import pythoncom
import win32com.client as win32

from yrx_project.scene.docs_processor.base import Command, ActionContext
from yrx_project.scene.docs_processor.command.n2n_position_command import SearchTextCommand, MoveCursorCommand
from yrx_project.scene.docs_processor.command.n2n_select_command import SelectCurrentCommand
from yrx_project.scene.docs_processor.command.n2n_update_command import ReplaceTextCommand
from yrx_project.scene.docs_processor.command.n2one_commands import MergeDocumentsCommand


class ActionParser:
    @staticmethod
    def parse(actions) -> typing.List[Command]:
        """
        [
            {"action_type": "position", "action": "find_first_after", "params": {"content": "123"}},
            {"action_type": "position", "action": "move_left", "params": {"content": "123"}},
            {"action_type": "position", "action": "move_down", "params": {"content": "123"}},
            {"action_type": "select", "action": "select_current_cell", "params": {"content": "123"}},
            {"action_type": "update", "action": "replace", "params": {"content": "123"}},
            {"action_type": "n2m", "action": "merge_docs", "params": {"inputs": [], "outputs": ""}},
        ]
        """
        commands_obj_list = []
        for action_dict in actions:
            action_type = action_dict.get("action_type")
            action = action_dict.get("action")
            params = action_dict.get("params", {})

            try:
                if action == "find_first_after":
                    commands_obj_list.append(
                        SearchTextCommand(
                            content=params.get("content"),
                            pointer_after_search="after",
                            action_type=action_type,
                            action_name=action,
                        )
                    )
                elif action in ("move_left", "move_down", "move_right", "move_up"):
                    direction = action.split('_')[1]
                    commands_obj_list.append(
                        MoveCursorCommand(
                            direction=direction,
                            units=params.get("units", 1),
                            action_type=action_type,
                            action_name=action,
                        )
                    )
                elif action.startswith("select_current_"):
                    current_type = action.split('_')[-1]
                    commands_obj_list.append(
                        SelectCurrentCommand(
                            current=current_type,
                            action_type=action_type,
                            action_name=action,
                        )
                    )
                elif action == "replace":
                    commands_obj_list.append(
                        ReplaceTextCommand(
                            content=params.get("content"),
                            action_type=action_type,
                            action_name=action,
                        )
                    )
                elif action == "merge_docs":
                    commands_obj_list.append(
                        MergeDocumentsCommand(
                            inputs=params.get("inputs", []),
                            output=params.get("output"),
                            action_type=action_type,
                            action_name=action,
                        )
                    )
            except KeyError as e:
                raise ValueError(f"Missing required parameter for action {action}: {e}")

        return commands_obj_list


class ActionProcessor:
    def __init__(self, config, after_each_file_func=None, after_each_action_func=None):
        self.commands = ActionParser.parse(config)
        self.context = ActionContext()
        self.merge_commands = [a for a in self.commands if a.action_type == "n2m"]
        self.normal_commands = [a for a in self.commands if a.action_type != "n2m"]

        self.after_each_file_func = after_each_file_func
        self.after_each_action_func = after_each_action_func

    def process(self, file_paths):
        pythoncom.CoInitialize()
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        self.context.word = word
        self.context.total_task = len(file_paths)
        try:
            # 处理普通文件操作
            for file_path in file_paths:
                if not os.path.exists(file_path):
                    raise FileNotFoundError(f"File not found: {file_path}")

                doc = word.Documents.Open(os.path.abspath(file_path))
                self.context.selection = word.Selection
                self.context.doc = doc
                self.context.file_path = file_path

                for command in self.normal_commands:
                    self.context.current_task = command.action_name
                    command.run(self.context)
                    if self.after_each_action_func is not None:
                        self.after_each_action_func(self.context)

                doc.Save()
                # doc.Close()
                self.context.doc = None
                self.context.selection = None

                self.context.done_task()  # todo: 应该是按照 map 和 reduce 进行阶段操作，而不是简单的统计文件的进度
                if self.after_each_file_func is not None:
                    self.after_each_file_func(self.context)

            # 处理合并操作
            self.context.msg = "正在处理聚合类操作..."
            for merge_action in self.merge_commands:
                merge_action.run(self.context)

        except Exception as e:
            print(f"Processing error: {str(e)}")
            raise
        finally:
            try:
                word.Quit()
            except:
                pass
            self.context.word = None
            pythoncom.CoUninitialize()

        return


if __name__ == '__main__':
    # ActionProcessor([
    #     {"action_type": "position", "action": "find_first_after", "params": {"content": "职务"}},
    #     # {"action_type": "position", "action": "move_left", "params": {"content": "123"}},
    #     {"action_type": "position", "action": "move_right", "params": {"content": 1}},
    #     {"action_type": "select", "action": "select_current_cell", "params": None},
    #     {"action_type": "update", "action": "replace", "params": {"content": "123abc123"}},
    #     # {"action_type": "n2m", "action": "merge_docs", "params": {"inputs": [], "outputs": ""}},
    # ]).process(file_paths=[
    #     r"D:\Projects\yrx_project\test.docx",
    # ])
    #


    ActionProcessor([
        # {'action': 'find_first_after', 'action_type': 'position', 'params': {'content': '教授', 'inputs': [r'D:/Projects/yrx_project/test.docx'], 'output': r'D:\Projects\yrx_project\tmp.docx'}},
        # {'action': 'move_down', 'action_type': 'position', 'params': {'content': '1', 'inputs': [r'D:/Projects/yrx_project/test.docx'], 'output': r'D:\Projects\yrx_project\tmp.docx'}},
        # {'action': 'select_current_cell', 'action_type': 'select', 'params': {'content': '', 'inputs': [r'D:/Projects/yrx_project/test.docx'], 'output': r'D:\Projects\yrx_project\tmp.docx'}},
        # {'action': 'replace', 'action_type': 'update', 'params': {'content': '撒扩大飞机阿萨', 'inputs': [r'D:/Projects/yrx_project/test.docx'], 'output': r'D:\Projects\yrx_project\tmp.docx'}}
        {'action_type': 'n2m', 'action': 'merge_docs', 'params': {'content': '撒扩大飞机阿萨', 'inputs': [
        r"D:\Projects\yrx_project\test1.docx",
        r"D:\Projects\yrx_project\test2.docx",
        r"D:\Projects\yrx_project\test3.docx",
    ], 'output': r'D:\Projects\yrx_project\tmp1.docx'}}
    ]).process(file_paths=[
        r"D:\Projects\yrx_project\test1.docx",
        r"D:\Projects\yrx_project\test2.docx",
        r"D:\Projects\yrx_project\test3.docx",
    ])