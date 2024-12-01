import ast

import pandas as pd

from yrx_project.scene.merged_cell.const import LanguageEnum
from yrx_project.utils.code_util import PythonCodeParser
from yrx_project.utils.df_util import MergedCells, generate_unique_column_name


# 代码提取功能：检查
def check_code(language: LanguageEnum, code_text: str) -> (bool, str):
    if "os." in code_text:
        return False, "⚠️存在危险代码"
    try:
        parsed_code = ast.parse(code_text)
        return True, ""
    except SyntaxError as e:
        error_details = (
            f"❌语法校验失败\n"
            f"错误信息: {e.msg}\n"
            f"行号: {e.lineno}\n"
            f"偏移（列）: {e.offset}\n"
            f"错误文本: {e.text.strip() if e.text else 'N/A'}"
        )
        return False, error_details


def do_with_code(language: LanguageEnum, code_text: str, df, merged_cells: MergedCells, ind):
    """
    :param language: 语言类型
    :param code_text: 代码内容：默认执行定义的apply函数
    :param df: DataFrame
    :param merged_cells: MergedCells对象
    :param ind: 操作的索引
    """
    """
def apply(name):
    print(name)
    """
    # 解析函数
    user_func = None
    if language == LanguageEnum.PYTHON:
        python_parser = PythonCodeParser(code_text=code_text, entry_func="apply")
        pass_check, error = python_parser.check_code()
        if not pass_check:
            return {"is_success": False, "error": error}

        user_func = python_parser.get_func()

    if user_func is None:
        return {"is_success": False, "error": "无法找到函数"}
    # 执行内容
    col_name = generate_unique_column_name(df, f"{language}处理结果")
    df.insert(ind+1, col_name, df.fillna("").apply(lambda x: user_func(x), axis=1))
    merged_cells.insert_col(ind)
    merged_cells.add_col_merged_cell_mapping(ind+1, copy_from=ind)
    df.merged_cells = merged_cells
    return {"is_success": True, "df": df}


def sort_merged_cell_df(df, merged_cells: MergedCells, col_index, rank_order):
    """
    对df的指定列，col_index列排序（按照给定顺序 rank_order）
    merged_cells: [(min_row, min_col, max_row, max_col), ()]
    """
    pass
    # # 步骤1：创建一个从 rank_order 到数值的映射，用于排序
    # rank_map = {value: i for i, value in enumerate(rank_order)}
    #
    # # 步骤2：创建一个字典来记录每个合并单元格的行索引
    # merge_map = {}
    # for (min_row, min_col, max_row, max_col) in merged_cells.iter(index=False):
    #     for row in range(min_row - 1, max_row):
    #         if row not in merge_map:
    #             merge_map[row] = list(range(min_row - 1, max_row))
    #
    # # 步骤3：创建一个新的 DataFrame 来存储排序后的结果
    # sorted_df = pd.DataFrame(columns=df.columns)
    #
    # # 步骤4：对 DataFrame 的指定列进行排序
    # df['_rank'] = df.iloc[:, col_index].map(rank_map)
    # sorted_indices = df.sort_values(by='_rank').index
    #
    # # 步骤5：根据排序后的索引和合并信息重新排列 DataFrame
    # visited = set()
    # for idx in sorted_indices:
    #     if idx not in visited:
    #         if idx in merge_map:
    #             # 获取合并单元格的所有行
    #             rows_to_move = merge_map[idx]
    #             # 将这些行添加到新的 DataFrame 中
    #             sorted_df = pd.concat([sorted_df, df.iloc[rows_to_move]], ignore_index=True)
    #             # 标记这些行为已访问
    #             visited.update(rows_to_move)
    #         else:
    #             # 如果不是合并单元格的一部分，直接添加
    #             sorted_df = pd.concat([sorted_df, df.iloc[[idx]]], ignore_index=True)
    #             visited.add(idx)
    #
    # # 步骤6：移除临时排序列
    # sorted_df.drop(columns=['_rank'], inplace=True)
    #
    # # 步骤7：调整合并单元格的信息
    # new_merged_cells = []
    # for (min_row, min_col, max_row, max_col) in merged_cells.iter(index=False):
    #     # 计算新的合并单元格范围
    #     original_indices = list(range(min_row - 1, max_row))
    #     new_indices = sorted_df.index[sorted_df.index.isin(original_indices)].tolist()
    #     if new_indices:
    #         new_min_row = min(new_indices) + 1
    #         new_max_row = max(new_indices) + 1
    #         new_merged_cells.append((new_min_row, min_col, new_max_row, max_col))
    #
    # sorted_df.new_merged_cells = MergedCells(new_merged_cells)
    # return sorted_df
    #
    # # return sorted_df, new_merged_cells
    #
    # # 步骤1：创建一个从 rank_order 到数值的映射，用于排序
    # rank_map = {value: i for i, value in enumerate(rank_order)}
    #
    # # 步骤2：在 DataFrame 中添加一个临时列用于排序
    # df['_rank'] = df.iloc[:, col_index].map(rank_map)
    #
    # # 步骤3：根据临时排序列对 DataFrame 进行排序
    # df = df.sort_values(by='_rank').reset_index(drop=True)
    #
    # # 步骤4：移除临时排序列
    # df.drop(columns=['_rank'], inplace=True)
    #
    # # 步骤5：调整合并单元格的信息
    # new_merged_cells = []
    # for (min_row, min_col, max_row, max_col) in merged_cells.iter(index=False):
    #     # 计算排序后合并单元格范围的新位置
    #     # 我们需要找到排序后 DataFrame 中原始行的新索引
    #     original_indices = list(range(min_row - 1, max_row))
    #     sorted_indices = df.index[df.index.isin(original_indices)].tolist()
    #
    #     # 计算排序后新的最小和最大行索引
    #     if sorted_indices:
    #         new_min_row = min(sorted_indices) + 1
    #         new_max_row = max(sorted_indices) + 1
    #         new_merged_cells.append((new_min_row, min_col, new_max_row, max_col))
    #
    # df.merged_cells = MergedCells(new_merged_cells)
    # return df
