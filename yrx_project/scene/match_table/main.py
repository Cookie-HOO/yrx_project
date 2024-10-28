import typing

import numpy as np
import pandas as pd


def read_table(excel_config: dict) -> pd.DataFrame:
    """
    :param excel_config:
    :param excel_config: [
        {
            "file_path": "./data/test.xlsx",
            "sheet_name_or_index": "Sheet1",  # sheet名称或索引
            "col_row": 2,  # 表头所在的行号（从1开始）
        }
    ]
    :return:
    """
    df = pd.read_excel(excel_config["file_path"], sheet_name=excel_config["sheet_name_or_index"], header=excel_config["col_row"] - 1)
    return df


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
    :return:
        pd.DataFrame:  拼接后的df
        np.array:  匹配到的行索引，如果有多组匹配，取最后一个匹配
        np.array: 未匹配的行索引，如果有多组匹配，取最后一个匹配
    """
    matched_indices = []
    unmatched_indices = []
    for match_dict in match_cols_and_df:
        matched_indices = []
        unmatched_indices = []

        match_df = match_dict['df']
        match_cols = match_dict['match_cols']
        catch_cols = match_dict['catch_cols']

        match_col_names = [i['match_col'] for i in match_cols]
        match_df = match_df.drop_duplicates(subset=match_col_names, keep='first')

        for col_dict in match_cols:
            main_col = col_dict['main_col']
            match_col = col_dict['match_col']
            matched_rows = main_df[main_col].isin(match_df[match_col])
            matched_indices.extend(main_df[matched_rows].index.values)
            unmatched_indices.extend(main_df[~matched_rows].index.values)
            main_df = pd.merge(main_df, match_df[catch_cols + [match_col]], how='left', left_on=main_col,
                               right_on=match_col)
    return main_df, np.array(matched_indices), np.array(unmatched_indices)


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
    df, matched_indices, unmatched_indices = match_table(df1, [
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

