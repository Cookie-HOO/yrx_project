import json
import os
from typing import List, Dict, Optional, Callable

import shutil
from multiprocessing import cpu_count, Pool

from PyQt5.QtWidgets import QApplication

from yrx_project.client.const import COLOR_YELLOW, DROPDOWN, READONLY_VALUE
from yrx_project.client.utils.table_widget import TableWidgetWrapper
from yrx_project.scene.process_docs.base import ActionContext, ACTION_TYPE_MAPPING
from yrx_project.scene.process_docs.action_types import action_types
from yrx_project.scene.process_docs.const import SCENE_TEMP_PATH
from yrx_project.scene.process_docs.processor import ActionProcessor
from yrx_project.utils.iter_util import swap_items_in_origin_list


current_suit = []

# 获取指定的doc的页数
def get_docx_pages(file_path):
    import win32com.client as win32
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Open(file_path)
        # 使用Word内置统计功能
        pages = doc.ComputeStatistics(2)  # 2 = wdNumberOfPages
        return pages
    finally:
        doc.Close()
        word.Quit()


def get_docx_pages_with_multiprocessing(file_paths):
    if len(file_paths) == 1:
        file_path = file_paths[0]
        return [get_docx_pages(file_path)]

    # 多于1个用多进程
    num_cores = cpu_count()  # 获取CPU的核心数
    with Pool(processes=min(num_cores, len(file_paths))) as pool:  # 创建一个包含4个进程的进程池
        results = pool.map(get_docx_pages, file_paths)  # 将函数和参数列表传递给进程池
    return results


class ActionRunner:
    def __init__(
        self,
        input_paths: List[str],
        df_actions,
        after_each_action_func: Optional[Callable[[ActionContext], None]] = None,
    ):
        self.input_paths = input_paths
        self.df_actions = df_actions
        self.after_each_action_func = after_each_action_func
        self.action_objs = self._prepare_action_objs()
        self.processor: Optional[ActionProcessor] = None

    def _prepare_action_objs(self) -> List[Dict]:
        """将DataFrame中的动作转换为对象列表"""
        action_objs = []
        for _, row in self.df_actions.iterrows():
            action_objs.append({
                "action_id": row["__动作id"],
                "action_content": row["动作内容"],
            })
        return action_objs

    def run_actions(self):
        """一次性执行所有动作"""
        ActionProcessor(
            self.action_objs,
            after_each_action_func=self.after_each_action_func
        ).process(self.input_paths)

    def debug_actions(self):
        """进入调试模式，初始化单步执行环境"""
        self.processor = ActionProcessor(
            self.action_objs,
            after_each_action_func=self.after_each_action_func,
            debug_mode=True
        )
        # 初始化上下文但不立即执行
        self.processor.init_context(self.input_paths)
        self.processor.process_next_or_init()

    def next_action_or_cleanup(self) -> bool:
        """执行下一步动作，返回是否还有后续动作
        最后需要再推一次进行清理
        """
        if not self.processor:
            raise RuntimeError("Debug mode not initialized. Call debug_actions() first.")
        return self.processor.process_next_or_init()



def build_action_types_menu(table_wrapper: TableWidgetWrapper):
    """将底层的action_type封装成可以做menu的格式
    [
        {"type": "menu", "name": "定位", "children": [
            {"type": "menu_action", "name": "选项1", "func": lambda: 1},
            {"type": "menu_action", "name": "选项2", "func": lambda: 1},
        ]},
        {"type": "menu_spliter"},
        {"type": "menu_action", "name": "选项2", "func": lambda: 1},
    ]

    ACTION_TYPE_MAPPING = {  # id -> name
        "locate": "定位",
        "select": "选择",
        "update": "修改",
        "mixing": "混合",
    }

    # columns = ["action_type_id", "group_id", "action_id", "action_name", "action_content_ui", "action_content_limit", "command_class", "command_init_kwargs"]
    df = action_types.action.types_df

    将df转成上面list的格式，其中
    1. 同样的action_type_name，放到一起，所有对应的action放到children中
    2. 所有的func 都是 lambda: 1
    3. 遍历df的过程中，遇到 action_type_id 相同，但是 group_id不同，那么插入一个 {"type": "menu_spliter"},
    """
    action_suit_table = ActionSuitTableTrans(table_wrapper)

    # 获取 DataFrame
    df = action_types.action_types_df

    # 初始化结果列表
    menu_list = []

    # 按 group_id 分组
    action_type_dfs = df.groupby("action_type_id", sort=False)
    for action_type_id, action_type_df in action_type_dfs:
        action_type_name, action_type_tip = ACTION_TYPE_MAPPING[action_type_id]
        children = []
        action_type_group_dfs = action_type_df.groupby("group_id", sort=False)
        for _, action_type_group_df in action_type_group_dfs:
            if children:
                children.append({
                    "type": "menu_spliter",
                })
            for _, row in action_type_group_df.iterrows():
                ui_type = row["action_content_ui"]
                value = row["action_content_value"]
                tip = row["action_content_tip"]
                action_name = row["action_name"]
                action_id = row["action_id"]
                children.append({
                    "type": "menu_action",
                    "name": action_name,
                    "tip": tip,
                    "func": lambda _, action_type_name=action_type_name, action_id=action_id, action_name=action_name,
                                   ui_type=ui_type, value=value: table_wrapper.add_rich_widget_row([
                        {
                            "type": "readonly_text",
                            "value": action_type_name,  # 类型
                        }, {
                            "type": "readonly_text",
                            "value": action_name,  # 动作
                        }, {
                            "type": ui_type,  # 动作内容
                            "value": value,
                        }, {
                            "type": "button_group",
                            "values": [
                                {
                                    "value": "⬆",
                                    "onclick": lambda row_index, col_index, row: action_suit_table.apply_suit(
                                        swap_items_in_origin_list(row_index, row_index-1, action_suit_table.trans2action_suit())
                                    ),
                                },
                                {
                                    "value": "⬇︎",
                                    "onclick": lambda row_index, col_index, row: action_suit_table.apply_suit(
                                        swap_items_in_origin_list(row_index, row_index-1, action_suit_table.trans2action_suit())
                                    ),
                                },
                                {
                                    "value": "❌",
                                    "onclick": lambda row_index, col_index, row_: table_wrapper.delete_row(
                                        row_index),
                                },
                            ],
                        }, {
                            "type": "readonly_text",  # __动作id
                            "value": action_id,
                        },
                    ])
                })
        menu_list.append({"type": "menu", "name": action_type_name, "tip": action_type_tip, "children": children})
    return menu_list


def build_action_suit_menu(table_wrapper: TableWidgetWrapper):
    """
    复制当前结构
    加载复制的结构，如果有内容，进行提示（无法找回当前的内容）
    加载内置的结构（当前的按钮）
    ---
    查看当前逻辑计划执行图（待商榷）
    """
    action_suit_table = ActionSuitTableTrans(table_wrapper)

    menu_list = [
        {
             "type": "menu", "name": "📂加载内置动作流",
             "children": [{
                "type": "menu_action",
                "name": action_suit_name,
                "func": lambda _, action_suit=actions: action_suit_table.apply_suit(action_suit),
                } for action_suit_name, actions in ACTION_SUITS.items()]
        },
        {
            "type": "menu_action", "name": "📂加载复制的动作流",
            "func": lambda _: action_suit_table.apply_from_clipboard(),
        },
        {
            "type": "menu_action", "name": "📄复制当前动作流",
            "func": lambda _: action_suit_table.copy2clipboard(),
        },

        # {
        #     "type": "menu_spliter",
        # },
        # {
        #     "type": "menu_action", "name": "🏃查看当前逻辑计划执行图",
        #     "func": lambda _: 1,  # todo
        # }
    ]

    return menu_list


class ActionSuitTableTrans:
    """table_wrapper 响应 action_suit的变化
    1. 将指定的action_suit 应用到当前table中（先清空，如果已经有值弹窗提示）
    2. 复制当前table中的suit 为json（剪切板）
    3. 剪切板中的suit 应用到当前table中
    """
    def __init__(self, table_wrapper: TableWidgetWrapper):
        self.table_wrapper = table_wrapper
        self.client_window = self.table_wrapper.table_widget.window()


    def _apply_suit(self, action_suit):
        """[{action_id: "", content: ""}, {}]"""
        self.table_wrapper.clear()
        for action in action_suit:
            if not action.get("action_id"):
                continue  # todo: 这是spliter
            action_row = action_types.get_action_by_id(action["action_id"])
            action_type_name, action_type_tip = ACTION_TYPE_MAPPING.get(action_row["action_type_id"])
            action_content_ui = action_row["action_content_ui"]

            # 下拉选择的话，需要特殊处理，值从action_type中来，设置修改默认值
            cur_index = None
            values = None
            if action_content_ui == DROPDOWN:
                values = action_row["action_content_value"]  # 来自 action_types，dropdown的话，就是范围限制
                cur_index = values.index(action["content"])

            action_ui, _, action_ui_options = action_row["action_content_ui"].partition("?")
            option_kv_list = [i.split("=") for i in action_ui_options.split(";")] if action_ui_options else []
            option_kv_dict = dict(option_kv_list) if option_kv_list else {}
            self.table_wrapper.add_rich_widget_row([
                {
                    "type": "readonly_text",
                    "value": action_type_name,  # 类型
                }, {
                    "type": "readonly_text",
                    "value": action_row["action_name"],  # 动作
                }, {
                    "type": action_ui,  # 动作ui和内容
                    "value": values or action.get("content") or READONLY_VALUE,
                    "cur_index": cur_index,
                    "options": option_kv_dict,
                }, {
                    "type": "button_group",
                    "values": [
                        {
                            "value": "⬆",
                            "onclick": lambda row_index, col_index, row: self.apply_suit(
                                swap_items_in_origin_list(row_index, row_index-1, self.trans2action_suit())
                            ),
                        },
                        {
                            "value": "⬇︎",
                            "onclick": lambda row_index, col_index, row: self.apply_suit(
                                swap_items_in_origin_list(row_index, row_index + 1, self.trans2action_suit())
                            ),
                        },
                        {
                            "value": "❌",
                            "onclick": lambda row_index, col_index, row_: self.table_wrapper.delete_row(
                                row_index),
                        },
                    ],
                }, {
                    "type": "readonly_text",  # __动作id
                    "value": action_row["action_id"],
                },

            ])

    def apply_suit(self, action_suit):
        # 冻结自动更新
        self.table_wrapper.table_widget.setUpdatesEnabled(False)
        self._apply_suit(action_suit)
        # 恢复更新
        self.table_wrapper.table_widget.setUpdatesEnabled(True)

    def apply_from_clipboard(self):
        action_suit_dumps = QApplication.clipboard().text()
        if not action_suit_dumps:
            self.client_window.modal(level="error", msg="❌剪切板为空")
            return
        try:
            action_suit = json.loads(action_suit_dumps)
        except Exception as e:
            self.client_window.modal(level="error", msg=f"❌剪切板内容解析失败: {action_suit_dumps}")
            return
        self.apply_suit(action_suit)
        self.client_window.modal(level="info", msg="✅已成功加载")

    def trans2action_suit(self) -> list:
        actions_df = self.table_wrapper.get_data_as_df()
        """__动作id, 动作内容"""
        actions_df["action_id"] = actions_df["__动作id"]
        actions_df["content"] = actions_df["动作内容"]
        action_suit_list = actions_df[["action_id", "content"]].to_dict(orient="records")
        return action_suit_list

    def copy2clipboard(self):
        action_suit_list = self.trans2action_suit()
        if not action_suit_list:
            self.client_window.modal(level="warn", msg="⚠️当前为空")
            return
        action_suit_dumps = json.dumps(action_suit_list, ensure_ascii=False)
        QApplication.clipboard().setText(action_suit_dumps)
        self.client_window.modal(level="info", msg="✅成功复制到剪贴板")


def cleanup_scene_folder():
    if os.path.exists(SCENE_TEMP_PATH):
        shutil.rmtree(SCENE_TEMP_PATH)
    os.makedirs(SCENE_TEMP_PATH)


def has_content_in_scene_folder():
    if os.path.exists(SCENE_TEMP_PATH):
        if os.listdir(SCENE_TEMP_PATH):
            return True
    return False


text = """


    该同志政治立场坚定，坚定“四个自信”，具有“四个意识”，做到“两个维护”，深刻认识“两个确立”的重大意义。热爱教师职业，肯于钻研，勤于总结，善于表达，工作踏实认真。遵纪守法，遵守学术规范和教师行为规范。








                    
单位党组织盖章：


                                  2025年2月24日"""
ACTION_SUITS = {
    "政审处理": [
        {"action_spliter": "1. 修改政审意见", "bg_color": COLOR_YELLOW},
        # 1. 搜索定位 姓名
        {"action_id": "search_first_and_move_right", "content": "民族"},
        # 2. 光标移动
        # {"action_id": "move_right_chars", "content": 1},
        {"action_id": "move_down_lines", "content": 3},
        # 3. 选择
        {"action_id": "select_current_scope", "content": "表格单元格"},
        # 4. 替换
        {"action_id": "replace_text", "content": text},

        {"action_spliter": "2. 修改第一部分字体", "bg_color": COLOR_YELLOW},
        # 5. 选择第一部分
        # {"action_id": "move_prev_to_landmark_only_text", "content": "当前单元格开头"},  # TODO 未测试通过
        # {"action_id": "select_current_scope", "content": "段落"},
        # 6. 修改字体
        {"action_id": "set_font_family", "content": "宋体"},
        # {"action_id": "set_font_color_custom", "content": "#000000"},
        {"action_id": "set_font_size", "content": "小四"},
        #
        # {"action_spliter": "3. 修改第二部分字体", "bg_color": COLOR_YELLOW},
        # # 7. 选中第二部分，设置字体
        # {"action_id": "select_next_to_landmark_only_text", "content": "当前单元格结尾"},
        # {"action_id": "set_font_family", "content": "宋体"},
        # {"action_id": "set_font_size", "content": "10pt"},

        # 9. 合并文档
        {"action_spliter": "4. 合并文档", "bg_color": COLOR_YELLOW},
        {"action_id": "merge_documents", "content": None},
    ]
}
