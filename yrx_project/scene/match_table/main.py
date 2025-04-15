import time
import typing

import pandas as pd

from yrx_project.scene.match_table.const import MATCH_OPTION, MAKEUP_MAIN_COL, ADD_COL_OPTION, MAKEUP_MAIN_COL_WITH_OVERWRITE
from yrx_project.utils.string_util import remove_by_ignore_policy

STR_EQUAL = "相等"
STR_CONTAINED = "被主表包含"

MATCH_FUNC_MAP = {
    STR_EQUAL: lambda m, h: m == h,
    STR_CONTAINED: lambda m, h: h in m,
}

def match_table(main_df, match_cols_and_df: typing.List[dict], add_overall_match_info=False) -> (pd.DataFrame, dict, dict):
    """
    :param main_df:
    :param add_overall_match_info:
        是否需要添加总体匹配信息
    :param match_cols_and_df:
        [
            {
                “id”: "",  #展示信息时，用这个id作为key
                "df": pd.DataFrame,
                "match_cols": [
                    {
                        "main_col": "a",
                        "match_col": "a",
                    },
                ],
                "catch_cols": [],  # 匹配到后，在辅助表中需要保留的列
                "match_func": lambda x, y: x == y,  # 匹配函数
                "match_ignore_policy":  # ["不忽略任何内容“]  或者  ["忽略所有中英文标点符号", "中文括号及内容"]
                "match_detail_text": lambda x, y: x == y,  # 匹配函数
                "match_detail_text":  # ｜ 分割的匹配到的，为匹配到的，为空的，额外展示的列
            }
        ]
    :return:
        pd.DataFrame:  拼接后的df
        overall_match_info: 总体匹配详情
            {
                "time_cost": 耗时,
                “match_index_list”: 匹配到的行索引,
                “unmatch_index_list”: 未匹配到的行索引,
                "no_content_index_list": 无内容的行索引,
            }
        detail_match_info：分辅助表的匹配详情

    匹配结果增加
        match_id%%{col}
        match_id%%匹配附加信息（文字）
        match_id%%匹配附加信息（行数）

    举例
        主表
            姓名 年龄
            张三 18
            张三 19
            李四 19
            王五 20
            赵六 21
        辅助表1
            姓名 年龄 身高
            张三 18 180
            李四 19 190
            张三 20 200
        辅助表2
            姓名 年龄 身高
            王五 18 180
            赵六 19 190
    结果
        姓名 年龄 辅助表1%%身高 辅助表1%%匹配附加信息（文字）   辅助表1%%匹配附加信息（行数） 辅助表2%%身高 辅助表2%%匹配附加信息（文字）   辅助表2%%匹配附加信息（行数） %任一条件匹配%  %全部条件匹配%
        张三 18 180\n200 "匹配到"  2  ""  "未匹配到"  0  是  否
        张三 19 180\n200 "匹配到"  2  ""  "未匹配到"  0  是  否
        李四 19 190 ""  1  ""  "未匹配到"  0  是  否
        王五 20 ""  "未匹配到"  0  180 ""  "匹配到"  1  是  否
        赵六 21 ""  "未匹配到"  0  190 ""  "匹配到"  1  是  否
    以上是将辅助表1和辅助表2的身高列，合并到主表中
    """
    # 返回的全局信息
    overall_match_info = {}
    # 返回的详细信息
    detail_match_info = {}

    # 全局的额外信息
    match_detail_text = []  # ["匹配到", "未匹配到"]
    no_content_indices = []
    match_mapping = {}  # {1: [1,3]}  主表的第一列，的第1和3cell 匹配到了
    first_match_text = []

    for match_dict in match_cols_and_df:  # 对应不同辅助表的多个条件
        start_for_one_df = time.time()

        # 1.获取变量
        match_id = match_dict["id"]  # 一般是辅助表的文件名
        match_df = match_dict['df']
        match_cols = match_dict['match_cols']  # [{"main_col": "", "match_col"}]
        catch_cols_with_policy = match_dict['catch_cols']  # [["a", "添加一列"], ["b", "补充到主表", "c"]]
        match_ignore_policy = match_dict['match_ignore_policy']  # ["不忽略任何内容“]  或者  ["忽略所有中英文标点符号", "中文括号及内容"]
        match_detail_text = match_dict['match_detail_text']  # 匹配到 ｜ 未匹配到 ｜ 无内容
        match_func = match_dict['match_func']  # lambda x, y: x == y

        # 2.变量校验
        ## 目前只支持单条件匹配
        col_dict = match_cols[0]  # [{"main_col": "", "match_col"}, ...]
        main_col = col_dict['main_col']
        match_col = col_dict['match_col']
        if main_col not in main_df.columns or match_col not in match_df.columns:
            continue

        ## catch cols 只获取存在的列
        catch_cols_with_policy = [i for i in catch_cols_with_policy if i[0] in match_df.columns]

        # 3. 核心逻辑
        ## 1. 根据match func 在match_df中找到匹配的行，可能有多行（可以先记录在main_df）中
        ## 2. 将标记为需要 「添加一列」的列，根据 找到的行 ，从match_df中取出来，拼接到main_df中（\n分割）
        ## 3. 将标记为需要 「补充到主表」的列，根据找到的行，补充到主表对应列中（\n分割）
        ##    .replace('', np.nan) 然后 再 fillna
        ## 4. 增加「匹配情况（文字）」列 和 「匹配情况（行数）」列
        striped_main_col = main_df[main_col].astype(str).apply(remove_by_ignore_policy, args=(match_ignore_policy,))
        striped_match_col = match_df[match_col].astype(str).apply(remove_by_ignore_policy, args=(match_ignore_policy,))


        match_detail_text = [i.strip() for i in match_detail_text.split("｜")]
        if not first_match_text:
            first_match_text = match_detail_text
        match_tip = match_detail_text[0] if len(match_detail_text) > 0 else ""
        unmatch_tip = match_detail_text[1] if len(match_detail_text) > 1 else ""
        no_content_tip = match_detail_text[2] if len(match_detail_text) > 2 else ""

        # None记录为main_df无内容，用于和未匹配到的[]区分
        # match_func 是一个自定义的纯函数，无法直接用merge（并非简单的等值判断）
        def match_text(row):
            if isinstance(row, list):
                if len(row) > 0:
                    return match_tip
                return unmatch_tip
            return no_content_tip
        main_df["%匹配行索引%"] = striped_main_col.apply(lambda row: [index for index, v in striped_match_col.items() if match_func(row, v)] if not pd.isnull(row) and row else None)
        main_df[f"{match_id}%%匹配附加信息（文字）"] = main_df["%匹配行索引%"].apply(match_text)
        main_df[f"{match_id}%%匹配附加信息（行数）"] = main_df["%匹配行索引%"].apply(len)
        match_extra_cols_index_list = [main_df.columns.get_loc(i) -1 for i in [f"{match_id}%%匹配附加信息（文字）", f"{match_id}%%匹配附加信息（行数）"]]

        # 携带列或者补充到主表
        catch_cols_index_list = []
        for catch_col_with_policy in catch_cols_with_policy:
            catch_col_name = catch_col_with_policy[0]
            if catch_col_with_policy[1] == ADD_COL_OPTION:  # 说明需要添加一列
                main_df[f'{match_id}%%{catch_col_name}'] = main_df['%匹配行索引%'].apply(
                    lambda indices: '\n'.join(
                        match_df.loc[indices, catch_col_name].astype(str).replace('nan', '')  # 处理 NaN
                    ) if (isinstance(indices, list) and indices) else ''
                )
                catch_cols_index_list.append(main_df.columns.get_loc(f'{match_id}%%{catch_col_name}') -1)

            elif catch_col_with_policy[1] == MAKEUP_MAIN_COL_WITH_OVERWRITE:
                main_col_name = catch_col_with_policy[2]
                main_df[main_col_name] = main_df.apply(
                    lambda row: '\n'.join(
                        match_df.loc[row['%匹配行索引%'], catch_col_name]
                        .astype(str)
                        .replace('nan', '')  # 处理 NaN
                    ) if (isinstance(row['%匹配行索引%'], list) and row['%匹配行索引%'])
                    else row[main_col_name] if row['%匹配行索引%'] == []  # 未匹配时保留原值
                    else '',  # 主表内容为空时设为空字符串
                    axis=1
                )
            elif catch_col_with_policy[1] == MAKEUP_MAIN_COL:
                main_col_name = catch_col_with_policy[2]
                main_df[main_col_name] = main_df.apply(
                    lambda row:
                    row[main_col_name] if row[main_col_name]
                    else '\n'.join(
                        match_df.loc[row['%匹配行索引%'], catch_col_name]
                        .astype(str)
                        .replace('nan', '')
                    ) if (
                            isinstance(row['%匹配行索引%'], list)
                            and row['%匹配行索引%']
                    )
                    else row[main_col_name] if (
                            isinstance(row['%匹配行索引%'], list)
                            and len(row['%匹配行索引%']) == 0
                    )
                    else '',
                    axis=1
                )


        # 拼接返回信息
        # 内容为空的行索引列表（原列值为 None）
        no_content_indices = main_df[main_df["%匹配行索引%"].isnull()].index
        # 匹配到的行索引列表（非空列表且长度>0）
        matched_mask = main_df["%匹配行索引%"].apply(
            lambda x: isinstance(x, list) and len(x) > 0
        )
        matched_indices = main_df[matched_mask].index
        # 未匹配到的行索引列表（非空列表但长度=0）
        unmatched_mask = main_df["%匹配行索引%"].apply(
            lambda x: isinstance(x, list) and len(x) == 0
        )
        unmatched_indices = main_df[unmatched_mask].index

        detail_match_info[match_id] = {
            "time_cost": time.time() - start_for_one_df,
            "match_index_list": matched_indices,
            "unmatch_index_list": unmatched_indices,
            "no_content_index_list": no_content_indices,
            "catch_cols_index_list": catch_cols_index_list,
            "match_extra_cols_index_list": match_extra_cols_index_list,
        }
        # 已经匹配的索引
        main_col_index = main_df.columns.get_loc(main_col) -1
        match_rows = match_mapping.get(main_col_index) or []
        # 加入新索引
        match_rows.extend(matched_indices)
        match_mapping[main_col_index] = match_rows
        # 清理临时列
        main_df.drop(columns=['%匹配行索引%'], inplace=True)


    # 组装总体信息
    total_length = len(main_df)
    sets = [set(value.get("match_index_list")) for value in detail_match_info.values()]
    union_set = set.union(*sets)
    intersection_set = set.intersection(*sets)

    union_set_present = round(len(union_set) / total_length * 100, 2)  # 任一匹配
    intersection_set_present = round(len(intersection_set) / total_length * 100, 2)  # 全部匹配
    overall_match_info["union_set_length"] = len(union_set)
    overall_match_info["intersection_set_length"] = len(intersection_set)
    overall_match_info["union_set_present"] = union_set_present
    overall_match_info["intersection_set_present"] = intersection_set_present
    overall_match_info["match_for_main_col"] = match_mapping  # key是主表的第几列，value是都有第几行匹配到了

    if add_overall_match_info:
        main_df["%任一条件匹配%"] = match_detail_text[1] if len(match_detail_text) > 1 else ""  # 默认设置为匹配不到
        main_df["%全部条件匹配%"] = match_detail_text[1] if len(match_detail_text) > 1 else ""  # 默认设置为匹配不到

        main_df.loc[list(union_set), "%任一条件匹配%"] = first_match_text[0] if len(first_match_text) > 0 else ""  # 匹配到
        main_df.loc[no_content_indices, "%任一条件匹配%"] = first_match_text[2] if len(first_match_text) > 2 else ""  # 空

        main_df.loc[list(intersection_set), "%全部条件匹配%"] = first_match_text[0] if len(first_match_text) > 0 else ""  # 匹配到
        main_df.loc[no_content_indices, "%全部条件匹配%"] = first_match_text[2] if len(first_match_text) > 2 else ""  # 空

        overall_match_info["match_extra_cols"] = [main_df.columns[-1], main_df.columns[-2]]
        overall_match_info["match_extra_cols_index_list"] = [len(main_df.columns)-1, len(main_df.columns)-2]

    return main_df, overall_match_info, detail_match_info


if __name__ == '__main__':
    file_path = "./北京大学境外教材选用目录（2024年8月更新）_副本.xlsx"
    sheet_name_or_index = 0
    col_row = 5

    df1 = pd.DataFrame({
        "a": [1, 2, 3, 3],
        "b": [4, 5, 6, 5],
        "c": [7, 8, 9, 10]
    })

    df2 = pd.DataFrame({
        "a": [1, 3, 3],
        "b": [4, 5, 6],
        "c": [7, 8, 9]
    })
    """
    # 1.获取变量
    match_id = match_dict["id"]  # 一般是辅助表的文件名
    match_df = match_dict['df']
    match_cols = match_dict['match_cols']  # [{"main_col": "", "match_col"}]
    catch_cols_with_policy = match_dict['catch_cols']  # [["a", "添加一列"], ["b", "补充到主表", "c"]]
    match_ignore_policy = match_dict['match_ignore_policy']  # ["不忽略任何内容“]  或者  ["忽略所有中英文标点符号", "中文括号及内容"]
    match_detail_text = match_dict['match_detail_text']  # 匹配到 ｜ 未匹配到 ｜ 无内容
    match_func = match_dict['match_func']  # lambda x, y: x == y

    # 2.变量校验
    ## 目前只支持单条件匹配
    col_dict = match_cols[0]  # [{"main_col": "", "match_col"}, ...]
    main_col = col_dict['main_col']
    match_col = col_dict['match_col']
    if main_col not in main_df.columns or match_col not in match_df.columns:
        continue

    """
    result_df, result_overall, result_detail = match_table(df1, [
        {
            "id": "abcd",
            "df": df2,
            "match_cols": [
                {
                    "main_col": "a",
                    "match_col": "a",
                },
            ],
            "catch_cols": [["b", "添加一列"], ["c", "补充到主表", "c"]],
            "match_ignore_policy": ["不忽略任何内容"],
            "delete_policy": [MATCH_OPTION],
            "match_detail_text": "匹配到｜未匹配到｜无内容",
            "match_func": lambda x, y: x == y,
        }
    ])
