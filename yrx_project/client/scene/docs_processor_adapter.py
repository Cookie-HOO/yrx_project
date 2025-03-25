from multiprocessing import cpu_count, Pool

from yrx_project.client.utils.table_widget import TableWidgetWrapper
from yrx_project.scene.docs_processor.base import ActionContext, ACTION_TYPE_MAPPING
from yrx_project.scene.docs_processor.action_types import action_types
from yrx_project.scene.docs_processor.processor import ActionProcessor


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


def run_with_actions(
        input_paths,
        df_actions,
        after_each_action_func: [[ActionContext], None]=None
):
    """
    df_actions: 是一个df
        类型：中文
        动作：中文
        动作内容

    """
    action_objs = []
    for index, row in df_actions.iterrows():
        action_id = row["__动作id"]
        action_content = row["动作内容"]

        action_objs.append({
            "action_id": action_id,
            "action_params": {"content": action_content}},
        )

    ActionProcessor(
        action_objs,
        after_each_action_func=after_each_action_func,
    ).process(input_paths)


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
    # 获取 DataFrame
    df = action_types.action_types_df

    # 初始化结果列表
    menu_list = []

    # 按 group_id 分组
    action_type_dfs = df.groupby("action_type_id", sort=False)
    for action_type_id, action_type_df in action_type_dfs:
        action_type_name = ACTION_TYPE_MAPPING[action_type_id]
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
                action_name = row["action_name"]
                action_id = row["action_id"]
                children.append({
                    "type": "menu_action",
                    "name": action_name,
                    "func": lambda _, action_type_name=action_type_name, action_id=action_id, action_name=action_name, ui_type=ui_type,value=value: table_wrapper.add_rich_widget_row([
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
                            "type": "readonly_text",  # __动作id
                            "value": action_id,
                        }, {
                            "type": "button_group",
                            "values": [
                                {
                                    "value": "向下",
                                    "onclick": lambda row_index, col_index, row: table_wrapper.swap_rows(
                                        row_index, row_index+1),
                                },
                                {
                                    "value": "向上",
                                    "onclick": lambda row_index, col_index, row: table_wrapper.swap_rows(
                                        row_index, row_index-1),
                                },
                                {
                                    "value": "删除",
                                    "onclick": lambda row_index, col_index, row_: table_wrapper.delete_row(
                                        row_index),
                                },
                            ],
                        }
                    ])
                })
        menu_list.append({"type": "menu", "name": action_type_name, "children": children})
    return menu_list

if __name__ == '__main__':
    ActionProcessor([
        {"action_id": "find_first_after", "params": {"content": "职务"}},
        # {"action_type": "position", "action": "move_left", "params": {"content": "123"}},
        {"action_id": "move_right", "params": {"content": 1}},
        {"action_id": "select_current_cell", "params": None},
        {"action_id": "replace", "params": {"content": "123abc123"}},
        # {"action_type": "n2m", "action": "merge_docs", "params": {"inputs": [], "outputs": ""}},
    ]).process(file_paths=[
        r"D:\Projects\yrx_project\test.docx",
    ])