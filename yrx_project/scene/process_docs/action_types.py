import typing

import pandas as pd

from yrx_project.client.const import READONLY_TEXT, EDITABLE_TEXT, DROPDOWN, COLOR_STR_RED, COLOR_STR_YELLOW, \
    COLOR_STR_GREEN, COLOR_STR_BLUE, READONLY_VALUE, EDITABLE_INT, EDITABLE_COLOR
from yrx_project.scene.process_docs.base import Command, MIXING_TYPE_ID, ACTION_TYPE_MAPPING

# from yrx_project.scene.process_docs.command.insert_commands import InsertSpecialCommand, InsertTextCommand
# from yrx_project.scene.process_docs.command.locate_commands import SearchTextCommand, MoveCursorCommand, \
#     MoveCursorUntilSpecialCommand
# from yrx_project.scene.process_docs.command.select_commands import SelectCurrentScopeCommand, SelectUntilCommand
# from yrx_project.scene.process_docs.command.update_command import ReplaceTextCommand, UpdateFontCommand, \
#     AdjustFontSizeCommand, UpdateFontColorCommand, UpdateParagraphCommand
# from yrx_project.scene.process_docs.command.mixing_commands import MergeDocumentsCommand

from yrx_project.scene.process_docs.office_word_command_impl.commands import *

class ActionType:
    columns = ["action_type_id", "group_id", "action_id", "action_name", "action_content_ui", "action_content_tip", "action_content_value", "command_class", "command_init_kwargs"]

    def __init__(self):
        # 初始化 DataFrame
        self.action_types_df = pd.DataFrame(columns=self.columns)

    def add_action(self, action_type_id, group_id, action_id, action_name, action_content_ui, action_content_tip, action_content_value, command_class, command_init_kwargs=None):
        """
        添加一个新的动作类型到 DataFrame 中
        :param action_type_id: 动作类型的唯一标识符
        :param action_id: 动作的 ID
        :param action_name: 动作的名称
        :param action_content_ui: 动作内容的 UI
        :param action_content_tip: 动作内容的提示
        :param action_content_value: 动作内容的值
        :param group_id: 动作所属的组 ID
        :param command_class: 动作的命令类
        :param command_init_kwargs: 命令类的初始化参数
        """
        new_row = pd.DataFrame(
            [[action_type_id, group_id, action_id, action_name, action_content_ui,action_content_tip,  action_content_value, command_class, command_init_kwargs or {}]],
            columns=self.columns,
        )
        self.action_types_df = pd.concat([self.action_types_df, new_row], ignore_index=True)

    def load_from_config(self, action_config: dict):
        """
        {
            "locate": {
                "group1": [
                    ["action_id", "action_name", action_content_ui, action_content_tip, action_content_value, command_class, command_init_kwargs]
                ]
                "group2": [],
                }
        }
        """
        for action_type_id, group_action_dict in action_config.items():
            for group_id, actions in group_action_dict.items():
                for action in actions:
                    # action_id, action_name, action_content_ui, action_content_tip, action_content_value, command_class, command_init_kwargs = action
                    self.add_action(action_type_id, group_id, *action)

    def get_action_by_id(self, action_id):
        """
        根据 action_id 查询动作类型
        :param action_id: 动作的 ID
        :return: 匹配的动作类型（DataFrame 行）
        """
        return self.action_types_df[self.action_types_df["action_id"] == action_id].iloc[0]

    def get_actions_by_group(self, group_id):
        """
        根据 group_id 查询属于某个组的所有动作类型
        :param group_id: 组 ID
        :return: 匹配的动作类型（DataFrame 行）
        """
        return self.action_types_df[self.action_types_df["group_id"] == group_id]

    def init_command(self, action_id, action_content) -> Command:
        action_row = self.get_action_by_id(action_id)
        cc: typing.Type[Command] = action_row["command_class"]
        cik = action_row["command_init_kwargs"] or {}
        action_type_id = action_row["action_type_id"]
        action_type_name, action_type_tip = ACTION_TYPE_MAPPING.get(action_type_id)
        action_id = action_row["action_id"]
        action_name = action_row["action_name"]
        return cc(
            **cik,
            action_type_id=action_type_id, action_type_name=action_type_name,
            action_id=action_id, action_name=action_name, action_content=action_content

        )


action_types = ActionType()
action_types.load_from_config({
    "locate": {
        "search": [
            ["search_first_and_select", "搜索：[输入] 关键词并选中第一个匹配项", EDITABLE_TEXT, "输入搜索内容", "", SearchTextCommand, None],
            ["search_first_and_move_left", "搜索：[输入] 关键词并移动光标到左侧", EDITABLE_TEXT, "输入搜索内容", "", SearchTextCommand, {"pointer_after_search": "left"}],
            ["search_first_and_move_right", "搜索：[输入] 关键词并移动光标到右侧", EDITABLE_TEXT, "输入搜索内容", "", SearchTextCommand, {"pointer_after_search": "right"}],
        ],
        "move": [
            ["move_up_lines", "移动：[输入] 上移行数", EDITABLE_INT+"?suffix= 行;min_num=1", "输入正整数（默认：1）", 1, MoveCursorCommand, {"direction": "up", "unit": "行"}],
            ["move_down_lines", "移动：[输入] 下移行数", EDITABLE_INT+"?suffix= 行;min_num=1", "输入正整数（默认：1）", 1, MoveCursorCommand, {"direction": "down", "unit": "行"}],
            ["move_left_chars", "移动：[输入] 左移字符数", EDITABLE_INT+"?suffix= 字符;min_num=1", "输入正整数（默认：1）", 1, MoveCursorCommand, {"direction": "left", "unit": "字符"}],
            ["move_right_chars", "移动：[输入] 右移字符数", EDITABLE_INT+"?suffix= 字符;min_num=1", "输入正整数（默认：1）", 1, MoveCursorCommand, {"direction": "right", "unit": "字符"}],
        ],
        "move_until": [
            ["move_prev_to_landmark_only_text", "移动：向前跳至 [选择] 标识（忽略开头结尾的空白）", DROPDOWN, "请选择标识类型，所有内容只考虑文字，忽略空白和空行", ["当前行开头", "上一段开头", "上一段结尾", "当前单元格开头", "当前页面开头", "当前文档开头"], MoveCursorUntilSpecialCommand, {"direction": "left", "ignore_blank": True}],
            ["move_next_to_landmark_only_text", "移动：向后跳至 [选择] 标识（忽略开头结尾的空白）", DROPDOWN, "请选择标识类型，所有内容只考虑文字，忽略空白和空行", ["当前行结尾", "下一段开头", "下一段结尾", "当前单元格结尾", "当前页面结尾", "当前文档结尾"], MoveCursorUntilSpecialCommand, {"direction": "right", "ignore_blank": True}],
            ["move_prev_to_landmark", "移动：向前跳至 [选择] 标识", DROPDOWN, "请选择标识类型", ["当前行开头", "上一段开头", "上一段结尾", "当前单元格开头", "当前页面开头", "当前文档开头"], MoveCursorUntilSpecialCommand, {"direction": "left", "ignore_blank": False}],
            ["move_next_to_landmark", "移动：向后跳至 [选择] 标识", DROPDOWN, "请选择标识类型", ["当前行结尾", "下一段开头", "下一段结尾", "当前单元格结尾", "当前页面结尾", "当前文档结尾"], MoveCursorUntilSpecialCommand, {"direction": "right", "ignore_blank": False}],
        ]
    },
    "insert": {
        "default": [
            ["insert_special_symbol", "插入：[选择] 特殊符号类型", DROPDOWN, "请选择符号类型", ["分页符", "换行符", "分节符", "制表符"], InsertSpecialCommand, None],
            ["insert_custom_text", "插入：[输入] 自定义文本内容", EDITABLE_TEXT, "直接输入要插入的文字（支持多行）", "", InsertTextCommand, None],
        ],
    },
    "select": {
        "current": [
            ["select_current_scope", "选择当前范围：[选择] 范围类型", DROPDOWN, "请选择范围类型", ["行", "段落", "表格单元格", "页面", "整个文档"], SelectCurrentScopeCommand, None],
        ],
        "move_until_and_select": [
            ["select_prev_to_landmark_only_text", "选择：向前至 [选择] 标识（忽略开头结尾的空白）", DROPDOWN, "请选择标识类型，所有内容只考虑文字，忽略空白和空行", ["当前行开头", "上一段开头", "上一段结尾", "当前单元格开头", "当前页面开头", "当前文档开头"], MoveCursorUntilSpecialCommand, {"direction": "left", "ignore_blank": True, "select": True}],
            ["select_next_to_landmark_only_text", "选择：向后至 [选择] 标识（忽略开头结尾的空白）", DROPDOWN, "请选择标识类型，所有内容只考虑文字，忽略空白和空行", ["当前行结尾", "下一段开头", "下一段结尾", "当前单元格结尾", "当前页面结尾", "当前文档结尾"], MoveCursorUntilSpecialCommand, {"direction": "right","ignore_blank": True, "select": True}],
            ["select_prev_to_landmark", "选择：向前至 [选择] 标识", DROPDOWN, "请选择标识类型", ["当前行开头", "上一段开头", "上一段结尾", "当前单元格开头", "当前页面开头", "当前文档开头"], MoveCursorUntilSpecialCommand, {"direction": "left", "ignore_blank": False, "select": True}],
            ["select_next_to_landmark", "选择：向后至 [选择] 标识", DROPDOWN, "请选择标识类型", ["当前行结尾", "下一段开头", "下一段结尾", "当前单元格结尾", "当前页面结尾", "当前文档结尾"], MoveCursorUntilSpecialCommand, {"direction": "right","ignore_blank": False, "select": True}],
        ],
        "move_and_select": [
            ["select_up_lines", "选择：[输入] 向上选择的行数（算当前行）", EDITABLE_INT+"?suffix= 行;min_num=1", "输入正整数（默认：1）", "1", MoveCursorCommand,{"direction": "up", "unit": "line", "select": True}],
            ["select_down_lines", "选择：[输入] 向下选择的行数（算当前行）", EDITABLE_INT+"?suffix= 行;min_num=1", "输入正整数（默认：1）", "1", MoveCursorCommand, {"direction": "down", "unit": "line", "select": True}],
            ["select_left_chars", "选择：[输入] 向左选中字符", EDITABLE_INT+"?suffix= 字符;min_num=1", "输入正整数（默认：1）", "1", MoveCursorCommand, {"direction": "left", "unit": "character", "select": True}],
            ["select_right_chars", "选择：[输入] 向右选中字符", EDITABLE_INT+"?suffix= 字符;min_num=1", "输入正整数（默认：1）", "1", MoveCursorCommand, {"direction": "right", "unit": "character", "select": True}],
        ],
        "inline": [
            ["inline_select_to_prev_text", "行内：向前选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "left", "scope": "inline", "until": "custom"}],
            ["inline_select_to_next_text", "行内：向后选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "right", "scope": "inline", "until": "custom"}],
        ],
        "cell": [
            ["cell_select_to_prev_landmark", "表格单元格内：向前选择至 [选择] 终止条件", DROPDOWN, "请选择终止条件", ["第一个空行"], SelectUntilCommand, {"direction": "left", "scope": "cell", "until": "preset"}],
            ["cell_select_to_next_landmark", "表格单元格内：向后选择至 [选择] 终止条件", DROPDOWN, "请选择终止条件", ["第一个空行"], SelectUntilCommand, {"direction": "right", "scope": "cell", "until": "preset"}],
            ["cell_select_to_prev_text", "表格单元格内：向前选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "left", "scope": "cell", "until": "custom"}],
            ["cell_select_to_next_text", "表格单元格内：向后选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "right", "scope": "cell", "until": "custom"}],
        ],
        "paragraph": [
            ["paragraph_select_to_prev_landmark", "段内：向前选择至 [选择] 文档标识", DROPDOWN, "请选择文档标识", ["第一个空行"], SelectUntilCommand, {"direction": "left", "scope": "paragraph", "until": "preset"}],
            ["paragraph_select_to_next_landmark", "段内：向后选择至 [选择] 文档标识", DROPDOWN, "请选择文档标识", ["第一个空行"], SelectUntilCommand, {"direction": "right", "scope": "paragraph", "until": "preset"}],
            ["paragraph_select_to_prev_text", "段内：向前选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "left", "scope": "paragraph", "until": "custom"}],
            ["paragraph_select_to_next_text", "段内：向后选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "right", "scope": "paragraph", "until": "custom"}],
        ],
        "page": [
            ["page_select_to_prev_landmark", "页内：向前选择至 [选择] 文档标识", DROPDOWN, "请选择文档标识", ["第一个空行"], SelectUntilCommand, {"direction": "left", "scope": "page", "until": "preset"}],
            ["page_select_to_next_landmark", "页内：向后选择至 [选择] 文档标识", DROPDOWN, "请选择文档标识", ["第一个空行"], SelectUntilCommand, {"direction": "right", "scope": "page", "until": "preset"}],
            ["page_select_to_prev_text", "页内：向前选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "left", "scope": "page", "until": "custom"}],
            ["page_select_to_next_text", "页内：向后选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "right", "scope": "page", "until": "custom"}],
        ],
        "doc": [
            ["doc_select_to_prev_landmark", "文档：向前选择至 [选择] 文档标识", DROPDOWN, "请选择文档标识", ["第一个空行"], SelectUntilCommand, {"direction": "left", "scope": "doc", "until": "preset"}],
            ["doc_select_to_next_landmark", "文档：向后选择至 [选择] 文档标识", DROPDOWN, "请选择文档标识", ["第一个空行"], SelectUntilCommand, {"direction": "right", "scope": "doc", "until": "preset"}],
            ["doc_select_to_prev_text", "文档：向前选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "left", "scope": "doc", "until": "custom"}],
            ["doc_select_to_next_text", "文档：向后选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "right", "scope": "doc", "until": "custom"}],
        ],
    },
    "update": {
        "text": [
            ["replace_text", "替换：[输入] 新文本内容", EDITABLE_TEXT, "输入替换后的完整内容", "", ReplaceTextCommand, None],
        ],
        "font": [
            ["set_font_color_preset", "颜色：[选择] 预设颜色", DROPDOWN, "请选择颜色", ["红色", "蓝色", "绿色"], UpdateFontColorCommand, {"color_mode": "preset"}],
            ["set_font_color_custom", "颜色：[输入] 自定义颜色", EDITABLE_COLOR, "请输入颜色", "#000000", UpdateFontColorCommand, {"color_mode": "custom"}],
            ["set_font_family", "字体：[选择] 字体系列", DROPDOWN, "请选择字体", ["宋体", "黑体", "仿宋", "楷体", "Times New Roman", "Arial"], UpdateFontCommand, {"attribute": "family"}],
            ["set_font_size", "字号：[选择] 预设字号", DROPDOWN, "请选择字号", ["初号", "一号", "小三", "10pt", "12pt", "14pt", "16pt"], UpdateFontCommand, {"attribute": "size"}],
            ["adjust_font_size", "字号调整：[选择] 增大或减小", DROPDOWN, "请选择操作", ["增大一级", "减小一级"], AdjustFontSizeCommand, {"step": 1}],
        ],
        "paragraph": [
            ["set_line_spacing_times", "行距：[输入] 行距倍数", DROPDOWN, "默认1倍", ["1"], UpdateParagraphCommand, {"attribute": "line_spacing", "line_spacing_type": "times"}],
            ["set_line_spacing_min", "行距：[输入] 最小磅值", DROPDOWN, "默认12磅", ["12"], UpdateParagraphCommand, {"attribute": "line_spacing", "line_spacing_type": "min_bounds"}],
            ["set_line_spacing_custom", "行距：[输入] 自定义磅值", EDITABLE_INT+"?suffix= 磅;min_num=1", "默认12磅", 12, UpdateParagraphCommand, {"attribute": "line_spacing", "line_spacing_type": "fix"}],
            ["set_paragraph_alignment", "对齐：[选择] 对齐方式", DROPDOWN, "请选择对齐模式", ["左对齐", "居中", "右对齐", "两端对齐"], UpdateParagraphCommand, {"attribute": "alignment"}],
        ]
    },
    MIXING_TYPE_ID: {
        "operations": [
            ["merge_documents", "合并文档（无需输入）", READONLY_TEXT, "合并所有文档", READONLY_VALUE, MergeDocumentsCommand, None],
        ]
    }
})
