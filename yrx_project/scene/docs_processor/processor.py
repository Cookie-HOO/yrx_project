import os
import typing

import pythoncom
import win32com.client as win32

from yrx_project.scene.docs_processor.base import Command, ActionContext
from yrx_project.scene.docs_processor.my_types import at


class ActionParser:
    @staticmethod
    def parse(actions) -> typing.List[Command]:
        """
        [
            {"action_id": "search_first_after", "action_params": {"content": "123"}},
            {"action_id": "move_left", "action_params": {"content": "123"}},
            {"action_id": "move_down", "action_params": {"content": "123"}},
            {"action_id": "select_current_cell", "action_params": {"content": "123"}},
            {"action_id": "replace_text", "action_params": {"content": "123"}},
            {"action_id": "merge_docs", "action_params": {"inputs": [], "outputs": ""}},
        ]
        """
        commands_obj_list = []
        for action_dict in actions:
            at.add_command2list(commands_obj_list, action_dict.get("action_id"), action_dict.get("action_params"))
        return commands_obj_list


class ActionProcessor:
    def __init__(self, config, after_each_file_func=None, after_each_action_func=None):
        self.commands = ActionParser.parse(config)
        self.context = ActionContext()
        self.mixing_commands = [a for a in self.commands if a.action_type_id == "mixing"]
        self.normal_commands = [a for a in self.commands if a.action_type_id != "mixing"]

        self.after_each_file_func = after_each_file_func
        self.after_each_action_func = after_each_action_func

    def process(self, file_paths, **kwargs):
        pythoncom.CoInitialize()
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        self.context.word = word
        self.context.input_paths = file_paths
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
            self.context.msg = "正在处理混合类操作..."
            for merge_action in self.mixing_commands:
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
        # {'action_id': 'search_first_after', 'action_params': {'content': '教授', 'inputs': [r'D:/Projects/yrx_project/test.docx'], 'output': r'D:\Projects\yrx_project\tmp.docx'}},
        # {'action_id': 'move_down', 'action_params': {'content': '1', 'inputs': [r'D:/Projects/yrx_project/test.docx'], 'output': r'D:\Projects\yrx_project\tmp.docx'}},
        # {'action_id': 'select_current_cell',  'action_params': {'content': '', 'inputs': [r'D:/Projects/yrx_project/test.docx'], 'output': r'D:\Projects\yrx_project\tmp.docx'}},
        # {'action_id': 'replace_text', 'action_params': {'content': '撒扩大飞机阿萨', 'inputs': [r'D:/Projects/yrx_project/test.docx'], 'output': r'D:\Projects\yrx_project\tmp.docx'}}
        {'action_id': 'merge_docs', 'action_params': {'content': '撒扩大飞机阿萨', 'inputs': [
        r"D:\Projects\yrx_project\test1.docx",
        r"D:\Projects\yrx_project\test2.docx",
        r"D:\Projects\yrx_project\test3.docx",
    ], 'output': r'D:\Projects\yrx_project\tmp1.docx'}}
    ]).process(file_paths=[
        r"D:\Projects\yrx_project\test1.docx",
        r"D:\Projects\yrx_project\test2.docx",
        r"D:\Projects\yrx_project\test3.docx",
    ])