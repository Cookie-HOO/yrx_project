import pandas as pd

from yrx_project.client.const import READONLY_TEXT, EDITABLE_TEXT, DROPDOWN, COLOR_STR_RED, COLOR_STR_YELLOW, \
    COLOR_STR_GREEN, COLOR_STR_BLUE
from yrx_project.scene.process_docs.base import Command, MIXING_TYPE_ID
from yrx_project.scene.process_docs.command.insert_commands import InsertSpecialCommand, InsertTextCommand
from yrx_project.scene.process_docs.command.locate_commands import SearchTextCommand, MoveCursorCommand
from yrx_project.scene.process_docs.command.select_commands import SelectCurrentScopeCommand, SelectRangeCommand, \
    SelectUntilCommand
from yrx_project.scene.process_docs.command.update_command import ReplaceTextCommand, UpdateFontCommand, \
    AdjustFontSizeCommand, UpdateFontColorCommand, UpdateParagraphCommand
from yrx_project.scene.process_docs.command.mixing_commands import MergeDocumentsCommand


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
        return self.action_types_df[self.action_types_df["action_id"] == action_id].iloc[0]

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
            ["search_first_and_select", "搜索：[输入] 关键词并选中第一个匹配项", EDITABLE_TEXT, "输入搜索内容", "", SearchTextCommand, None],
            ["search_first_and_move_left", "搜索：[输入] 关键词并移动光标到左侧", EDITABLE_TEXT, "输入搜索内容", "", SearchTextCommand, {"pointer_after_search": "left"}],
            ["search_first_and_move_right", "搜索：[输入] 关键词并移动光标到右侧", EDITABLE_TEXT, "输入搜索内容", "", SearchTextCommand, {"pointer_after_search": "right"}],
        ],
        "move": [
            ["move_up_lines", "移动：[输入] 上移行数", EDITABLE_TEXT, "输入正整数（默认：1）", "1", MoveCursorCommand, {"direction": "up", "unit": "line"}],
            ["move_down_lines", "移动：[输入] 下移行数", EDITABLE_TEXT, "输入正整数（默认：1）", "1", MoveCursorCommand, {"direction": "down", "unit": "line"}],
            ["move_left_chars", "移动：[输入] 左移字符数", EDITABLE_TEXT, "输入正整数（默认：1）", "1", MoveCursorCommand, {"direction": "left", "unit": "character"}],
            ["move_right_chars", "移动：[输入] 右移字符数", EDITABLE_TEXT, "输入正整数（默认：1）", "1", MoveCursorCommand, {"direction": "right", "unit": "character"}],
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
        "inline": [
            ["inline_select_to_start", "行内：跳转到行首（无需输入）", READONLY_TEXT, "操作立即生效", "---", SelectRangeCommand, {"boundary": "line_start"}],
            ["inline_select_to_end", "行内：跳转到行尾（无需输入）", READONLY_TEXT, "操作立即生效", "---", SelectRangeCommand, {"boundary": "line_end"}],
            ["inline_select_to_prev_text", "行内：向前选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "left", "scope": "inline", "until": "custom"}],
            ["inline_select_to_next_text", "行内：向后选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "right", "scope": "inline", "until": "custom"}],
        ],
        "cell": [
            ["cell_select_to_start", "单元格：跳转到起始位置（无需输入）", READONLY_TEXT, "操作立即生效", "---", SelectRangeCommand, {"boundary": "cell_start"}],
            ["cell_select_to_end", "单元格：跳转到结束位置（无需输入）", READONLY_TEXT, "操作立即生效", "---", SelectRangeCommand, {"boundary": "cell_end"}],
            ["cell_select_to_prev_condition", "单元格：向前选择至 [选择] 终止条件", DROPDOWN, "请选择终止条件", ["第一个空行"], SelectUntilCommand, {"direction": "left", "scope": "cell", "until": "preset"}],
            ["cell_select_to_next_condition", "单元格：向后选择至 [选择] 终止条件", DROPDOWN, "请选择终止条件", ["第一个空行"], SelectUntilCommand, {"direction": "right", "scope": "cell", "until": "preset"}],
            ["cell_select_to_prev_text", "单元格：向前选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "left", "scope": "cell", "until": "custom"}],
            ["cell_select_to_next_text", "单元格：向后选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "right", "scope": "cell", "until": "custom"}],
        ],
        "page": [
            ["page_select_to_start", "页内：跳转至页首（无需输入）", READONLY_TEXT, "操作立即生效", "---", SelectRangeCommand, {"boundary": "page_start"}],
            ["page_select_to_end", "页内：跳转至页尾（无需输入）", READONLY_TEXT, "操作立即生效", "---", SelectRangeCommand, {"boundary": "page_end"}],
            ["page_select_to_prev_landmark", "页内：向前选择至 [选择] 文档标识", DROPDOWN, "请选择文档标识", ["页眉", "页脚", "分页符"], SelectUntilCommand, {"direction": "left", "scope": "page", "until": "preset"}],
            ["page_select_to_next_landmark", "页内：向后选择至 [选择] 文档标识", DROPDOWN, "请选择文档标识", ["页眉", "页脚", "分页符"], SelectUntilCommand, {"direction": "right", "scope": "page", "until": "preset"}],
            ["page_select_to_prev_text", "页内：向前选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "left", "scope": "page", "until": "custom"}],
            ["page_select_to_next_text", "页内：向后选择至 [输入] 终止文本", EDITABLE_TEXT, "输入终止文本", "", SelectUntilCommand, {"direction": "right", "scope": "page", "until": "custom"}],
        ],
        "doc": [
            ["doc_select_to_start", "文档：跳转至开头（无需输入）", READONLY_TEXT, "操作立即生效", "---", SelectRangeCommand, {"boundary": "doc_start"}],
            ["doc_select_to_end", "文档：跳转至结尾（无需输入）", READONLY_TEXT, "操作立即生效", "---", SelectRangeCommand, {"boundary": "doc_end"}],
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
            ["set_font_family", "字体：[选择] 字体系列", DROPDOWN, "请选择字体", ["宋体", "黑体", "仿宋", "楷体", "Times New Roman", "Arial"], UpdateFontCommand, {"attribute": "family"}],
            ["set_font_size", "字号：[选择] 预设字号", DROPDOWN, "请选择字号", ["初号", "一号", "小三", "12pt", "14pt", "16pt"], UpdateFontCommand, {"attribute": "size"}],
            ["adjust_font_size", "字号调整：[选择] 调整方式", DROPDOWN, "请选择操作", ["增大一级", "减小一级"], AdjustFontSizeCommand, {"step": 1}],
            ["set_font_color", "颜色：[选择] 预设颜色", DROPDOWN, "请选择颜色", ["红色", "蓝色", "绿色"], UpdateFontColorCommand, {"color_mode": "preset"}],
        ],
        "paragraph": [
            ["set_line_spacing", "行距：[输入] 行距倍数", DROPDOWN, "默认1倍", "1", UpdateParagraphCommand, {"attribute": "line_spacing", "line_spacing_type": "times"}],
            ["set_line_spacing", "行距：[输入] 最小磅值", DROPDOWN, "默认12磅", "12", UpdateParagraphCommand, {"attribute": "line_spacing", "line_spacing_type": "min_bounds"}],
            ["set_custom_line_spacing", "行距：[输入] 自定义磅值", EDITABLE_TEXT, "默认12磅", "12", UpdateParagraphCommand, {"attribute": "line_spacing", "line_spacing_type": "fix"}],
            ["set_paragraph_alignment", "对齐：[选择] 对齐方式", DROPDOWN, "请选择对齐模式", ["左对齐", "居中", "右对齐", "两端对齐"], UpdateParagraphCommand, {"attribute": "alignment"}],
        ]
    },
    MIXING_TYPE_ID: {
        "operations": [
            ["merge_documents", "合并文档（无需输入）", READONLY_TEXT, "合并所有文档", "---", MergeDocumentsCommand, None],
        ]
    }
})