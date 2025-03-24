import typing
from multiprocessing import cpu_count, Pool

import win32com.client as win32

from yrx_project.scene.docs_processor.base import ActionContext
from yrx_project.scene.docs_processor.const import ACTION_MAPPING
from yrx_project.scene.docs_processor.processor import ActionProcessor


# 获取指定的doc的页数
def get_docx_pages(file_path):
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
        output_path,  # 考虑一个路径，目前传递一个文件
        after_each_file_func: typing.Callable[[ActionContext], None]=None,
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
        action_type = row["类型"]
        action = row["动作"]
        action_content = row["动作内容"]

        action_objs.append({
            "action_type": ACTION_MAPPING.get(action_type).get("id"),
            "action": [i.get("id") for i in ACTION_MAPPING.get(action_type).get("children") if i.get("name") == action][0],
            "params": {"content": action_content, "inputs": input_paths, "output": output_path}},
        )

    ActionProcessor(
        action_objs,
        after_each_action_func=after_each_action_func,
        after_each_file_func=after_each_file_func
    ).process(input_paths)

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