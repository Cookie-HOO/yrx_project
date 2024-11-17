import time
import typing

import numpy as np
import pandas as pd

from yrx_project.scene.match_table.const import MATCH_OPTION, UNMATCH_OPTION, NO_CONTENT_OPTION
from yrx_project.utils.iter_util import dedup_list
from yrx_project.utils.string_util import remove_punctuation_and_spaces


def check_match_table(main_df, match_cols_and_df: typing.List[dict]) -> typing.List[dict]:
    """根据策略进行匹配检测，如果匹配过程中待匹配表中的匹配列有重复，导致最终匹配的结果行数增加，那么找到这个列，和其对应的重复值
        :param main_df:
        :param match_cols_and_df:
            [
                {
                    "df": pd.DataFrame,
                    "match_cols": [
                        {
                            "main_col": "a",
                            "match_col": "a",
                        },
                    ],
                    "catch_cols": [],  # 匹配到后，在辅助表中需要保留的列
                    "match_color": "red",
                    "unmatch_color": "green",
                    "match_policy": "last" | "first"  # 如果辅助表中出现重复，取第一个还是最后一个
                }
            ]
        :return: 按照当前规则匹配后，因为匹配表中的某列有重复，会返回重复的列名和对应的值
            [
                {
                    "duplicate_cols": {
                        "col_name": "a",
                        "cell_values": [],
                    }
                }
            ]
    """
    duplicate_info = []
    for match_dict in match_cols_and_df:
        match_df = match_dict['df']
        match_cols = match_dict['match_cols']
        for col_dict in match_cols:
            main_col = col_dict['main_col']
            match_col = col_dict['match_col']

            target_rows = match_df[match_df[match_col].isin(main_df[main_col])][match_col]
            # 检查这些行是否有重复
            duplicate_rows = target_rows.duplicated(keep=False)

            # 返回重复的行
            duplicate_values = target_rows[duplicate_rows].unique()
            if len(duplicate_values) > 0:
                duplicate_info.append({
                    "duplicate_cols": {
                        "col_name": match_col,
                        "cell_values": duplicate_values.tolist(),
                    }
                })
    return duplicate_info


def match_table(main_df, match_cols_and_df: typing.List[dict]) -> (pd.DataFrame, np.array, np.array):
    """
    :param main_df:
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
                "match_color": "red",
                "unmatch_color": "green",
                "match_policy": "last" | "first"  # 如果辅助表中出现重复，取第一个还是最后一个
                “delete_policy”: [],
                "match_detail_test":  # ｜ 分割的匹配到的，为匹配到的，为空的，额外展示的列
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
                "delete_index_list": 被删除的行索引,
            }
        detail_match_info：分辅助表的匹配详情

    匹配结果增加
        match_id\n{col}
        match_id\n匹配附加信息（文字）
        match_id\n匹配附加信息（行数）
    """
    detail_match_info = {}

    start_for_all_df = time.time()
    for match_dict in match_cols_and_df:
        start_for_one_df = time.time()

        # 获取变量
        match_id = match_dict["id"]
        match_df = match_dict['df']
        match_cols = match_dict['match_cols']  # [{"main_col": "", "match_col"}]
        catch_cols_ = match_dict['catch_cols']  # ["a", "b"]
        delete_policy = match_dict['delete_policy']  # ["a", "b"]
        match_detail_text = match_dict['match_detail_text']  # ["a", "b"]

        # 定义列名
        match_num_col_name = f"{match_id}%%匹配附加信息（行数）"
        match_text_col_name = f"{match_id}%%匹配附加信息（文字）"
        catch_cols = [f"{match_id}%%{i}" for i in catch_cols_ if i in match_df.columns]  # 只要找的着的
        match_df.rename(columns=dict(zip(catch_cols_, catch_cols)), inplace=True)

        match_col_names = [i['match_col'] for i in match_cols]
        match_df[match_num_col_name] = match_df.groupby(match_col_names)[match_col_names[0]].transform('size')
        match_df = match_df.drop_duplicates(subset=match_col_names, keep=match_dict["match_policy"])  # 先对match col列去重

        # 目前只支持单条件匹配
        col_dict = match_cols[0]
        main_col = col_dict['main_col']
        match_col = col_dict['match_col']
        striped_main_col = main_df[main_col].apply(lambda row: remove_punctuation_and_spaces(row))
        striped_match_col = match_df[match_col].apply(lambda row: remove_punctuation_and_spaces(row))
        matched_rows = striped_main_col.isin(striped_match_col)  # 一堆 True False

        # 寻找拼配、未匹配的索引
        matched_indices = main_df[matched_rows].index.values
        unmatched_indices = main_df[~matched_rows].index.values
        no_content_indices = striped_main_col[striped_main_col == ""].index.values

        # 匹配
        main_df["__主表匹配列"] = striped_main_col
        match_df["__辅助表匹配列"] = striped_match_col
        main_df = pd.merge(
            main_df, match_df[dedup_list(catch_cols + ["__辅助表匹配列", match_num_col_name])], how='left', left_on="__主表匹配列",
            right_on="__辅助表匹配列", suffixes=('', '_来自辅助表')
        )
        main_df = main_df.drop(columns=["__主表匹配列", "__辅助表匹配列"], axis=1)

        # 增加匹配情况（文字）
        match_detail_text = [i.strip() for i in match_detail_text.split("｜")]
        main_df[match_text_col_name] = ""
        main_df.loc[matched_indices, match_text_col_name] = match_detail_text[0] if len(match_detail_text) > 0 else ""
        main_df.loc[unmatched_indices, match_text_col_name] = match_detail_text[1] if len(match_detail_text) > 1 else ""
        main_df.loc[no_content_indices, match_text_col_name] = match_detail_text[2] if len(match_detail_text) > 2 else ""

        # 如果需要删除
        delete_index = set()
        for delete_p in delete_policy:
            if delete_p == MATCH_OPTION:
                delete_index.update(matched_indices)
            elif delete_p == UNMATCH_OPTION:
                delete_index.update(unmatched_indices)
            elif delete_p == NO_CONTENT_OPTION:
                delete_index.update(no_content_indices)

        if delete_index:
            main_df = main_df.drop(list(delete_index))

        detail_match_info[match_id] = {
            "time_cost": time.time() - start_for_one_df,
            "match_index_list": matched_indices,
            "unmatch_index_list": unmatched_indices,
            "no_content_index_list": no_content_indices,
            "delete_index_list": list(delete_index),

            "catch_cols": catch_cols,
            "match_extra_cols": [match_num_col_name, match_text_col_name],

            "catch_cols_index_list": [],
            "match_extra_cols_index_list": [],
        }

    for match_id, detail_match_info_one in detail_match_info.items():
        detail_match_info_one["catch_cols_index_list"] = [main_df.columns.get_loc(i) for i in detail_match_info_one["catch_cols"]]
        detail_match_info_one["match_extra_cols_index_list"] = [main_df.columns.get_loc(i) for i in detail_match_info_one["match_extra_cols"]]

    return main_df, detail_match_info


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
    df, detail_match_info = match_table(df1, [
        {
            "id": "abcd",
            "df": df2,
            "match_cols": [
                {
                    "main_col": "a",
                    "match_col": "a",
                },
            ],
            "catch_cols": ["b"],  # 匹配到后，在辅助表中需要保留的列
            "delete_policy": [MATCH_OPTION],
            "match_detail_text": "匹配到|未匹配到|无内容",
            "match_policy": "first",
        }
    ])

    res = check_match_table(df1, [
        {
            "df": df2,
            "match_cols": [
                {
                    "main_col": "a",
                    "match_col": "a",
                },
            ],
            "catch_cols": ["b"],  # 匹配到后，在辅助表中需要保留的列
            "match_color": "red",
            "unmatch_color": "green",
        }
    ])
    print()

