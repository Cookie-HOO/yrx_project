import pandas as pd

from yrx_project.client.const import READONLY_TEXT, EDITABLE_TEXT
from yrx_project.scene.docs_processor.base import Command, MIXING_TYPE_ID
from yrx_project.scene.docs_processor.command.n2n_position_command import SearchTextCommand, MoveCursorCommand
from yrx_project.scene.docs_processor.command.n2n_select_command import SelectCurrentCommand
from yrx_project.scene.docs_processor.command.n2n_update_command import ReplaceTextCommand
from yrx_project.scene.docs_processor.command.n2one_commands import MergeDocumentsCommand


class ActionType:
    columns = ["action_type_id", "group_id", "action_id", "action_name", "action_content_ui", "action_content_value", "command_class", "command_init_kwargs"]

    def __init__(self):
        # 初始化 DataFrame
        self.action_types_df = pd.DataFrame(columns=self.columns)

    def add_action(self, action_type_id, group_id, action_id, action_name, action_content_ui, action_content_value, command_class, command_init_kwargs=None):
        """
        添加一个新的动作类型到 DataFrame 中
        :param action_type_id: 动作类型的唯一标识符
        :param action_id: 动作的 ID
        :param action_name: 动作的名称
        :param action_content_ui: 动作内容的 UI
        :param action_content_value: 动作内容的值
        :param group_id: 动作所属的组 ID
        :param command_class: 动作的命令类
        :param command_init_kwargs: 命令类的初始化参数
        """
        new_row = pd.DataFrame(
            [[action_type_id, group_id, action_id, action_name, action_content_ui, action_content_value, command_class, command_init_kwargs or {}]],
            columns=self.columns,
        )
        self.action_types_df = pd.concat([self.action_types_df, new_row], ignore_index=True)

    def load_from_config(self, action_config: dict):
        """
        {
            "locate": {
                "group1": [
                    ["action_id", "action_name", action_content_ui, action_content_value, command_class, command_init_kwargs]
                ]
                "group2": [],
                }
        }
        """
        for action_type_id, group_action_dict in action_config.items():
            for group_id, actions in group_action_dict.items():
                for action in actions:
                    # action_id, action_name, action_content_ui, action_content_value, command_class, command_init_kwargs = action
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

    def init_command(self, action_id, action_params) -> Command:
        action_row = self.get_action_by_id(action_id)
        cc = action_row["command_class"]
        cik = action_row["command_init_kwargs"] or {}
        action_type_id = action_row["action_type_id"]
        action_name = action_row["action_name"]
        action_params = action_params or {}
        return cc(**cik, **action_params, action_type_id=action_type_id, action_name=action_name)


action_types = ActionType()
action_types.load_from_config({
    "locate": {
        "search": [
            ["search_first_after", "搜索后光标后移", EDITABLE_TEXT, "搜索内容", SearchTextCommand, {"pointer_after_search": "after"}],
        ],
        "move": [
            ["move_down", "光标下移", EDITABLE_TEXT, "1", MoveCursorCommand, {"direction": "down"}],
        ]
    },
    "select": {
        "current": [
            ["select_current_cell", "选择当前单元格", READONLY_TEXT, "---", SelectCurrentCommand, {"current": "cell"}],
        ]
    },
    "update": {
        "text": [
            ["replace_text", "替换文本", EDITABLE_TEXT, "", ReplaceTextCommand, None],
        ]
    },

    # 混合类动作
    MIXING_TYPE_ID: {
        "merge": [
            ["merge_docs", "合并所有文档", READONLY_TEXT, "---", MergeDocumentsCommand, None],
        ]
    }
})
