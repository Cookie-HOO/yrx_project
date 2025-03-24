import typing

import pandas as pd

from yrx_project.scene.docs_processor.base import Command
from yrx_project.scene.docs_processor.command.n2n_position_command import SearchTextCommand, MoveCursorCommand
from yrx_project.scene.docs_processor.command.n2n_select_command import SelectCurrentCommand
from yrx_project.scene.docs_processor.command.n2n_update_command import ReplaceTextCommand
from yrx_project.scene.docs_processor.command.n2one_commands import MergeDocumentsCommand


ACTION_TYPE_MAPPING = {
    "locate": "定位",
    "select": "选择",
    "update": "修改",
    "mixing": "混合",
}


class ActionType:
    columns = ["action_type_id", "group_id", "action_id", "action_name", "action_content_ui", "action_content_limit", "command_class", "command_init_kwargs"]

    def __init__(self):
        # 初始化 DataFrame
        self.action_types_df = pd.DataFrame(columns=self.columns)

    def add_action(self, action_type_id, group_id, action_id, action_name, action_content_ui, action_content_limit, command_class, command_init_kwargs=None):
        """
        添加一个新的动作类型到 DataFrame 中
        :param action_type_id: 动作类型的唯一标识符
        :param action_id: 动作的 ID
        :param action_name: 动作的名称
        :param action_content_ui: 动作内容的 UI
        :param action_content_limit: 动作内容的限制，是一个函数，必须为True才可以
        :param group_id: 动作所属的组 ID
        :param command_class: 动作的命令类
        :param command_init_kwargs: 命令类的初始化参数
        """
        new_row = pd.DataFrame(
            [[action_type_id, group_id, action_id, action_name, action_content_ui, action_content_limit, command_class, command_init_kwargs or {}]],
            columns=self.columns,
        )
        self.action_types_df = pd.concat([self.action_types_df, new_row], ignore_index=True)

    def load_from_config(self, action_config: dict):
        """
        {
            "locate": {
                "group1": [
                    ["action_id", "action_name", action_content_ui, action_content_limit, command_class, command_init_kwargs]
                ]
                "group2": [],
                }
        }
        """
        for action_type_id, group_action_dict in action_config.items():
            for group_id, actions in group_action_dict.items():
                for action in actions:
                    # action_id, action_name, action_content_ui, action_content_limit, command_class, command_init_kwargs = action
                    self.add_action(action_type_id, group_id, *action)

    def get_action_by_id(self, action_id):
        """
        根据 action_id 查询动作类型
        :param action_id: 动作的 ID
        :return: 匹配的动作类型（DataFrame 行）
        """
        return self.action_types_df[self.action_types_df["action_id"] == action_id]

    def get_actions_by_group(self, group_id):
        """
        根据 group_id 查询属于某个组的所有动作类型
        :param group_id: 组 ID
        :return: 匹配的动作类型（DataFrame 行）
        """
        return self.action_types_df[self.action_types_df["group_id"] == group_id]

    def add_command2list(self, command_list: typing.List[Command], action_id, action_params):
        action_row = self.get_action_by_id(action_id)
        cc = action_row["command_class"]
        cik = action_row["command_init_kwargs"] or {}
        action_type_id = action_row["action_type_id"]
        action_name = action_row["action_name"]
        command_list.append(cc(**cik, **action_params, action_type_id=action_type_id, action_name=action_name))


at = ActionType()
at.load_from_config({
    "locate": {
        "search": [
            ["search_first_after", "搜索后光标后移", "editable_text", lambda x: len(str(x)) > 0, SearchTextCommand, {"pointer_after_search": "after"}],
        ],
        "move": [
            ["move_down", "光标下移", "editable_text", lambda x: x.isdigit(), MoveCursorCommand, {"direction": "down"}],
        ]
    },
    "select": {
        "current": [
            ["select_current_cell", "选择当前单元格", None, None, SelectCurrentCommand, {"current": "cell"}],
        ]
    },
    "update": {
        "text": [
            ["replace_text", "替换文本", "editable_text", lambda x: len(str(x)) > 0, ReplaceTextCommand, None],
        ]
    },
    "mixing": {
        "merge": [
            ["merge_docs", "合并所有文档", None, None, MergeDocumentsCommand, None],
        ]
    }
})
