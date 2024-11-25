import os
import typing
from functools import lru_cache
from multiprocessing import Pool, cpu_count, Manager

import pandas as pd
import xlrd
from openpyxl.reader.excel import load_workbook

# d = Manager().dict()  # 创建一个可以在多个进程之间共享的字典


def is_empty(value):
    if pd.isna(value) or value is None:
        return True
    return len(str(value)) == 0


def is_not_empty(value):
    return not is_empty(value)


def is_any_empty(*values):
    for i in values:
        if is_empty(i):
            return True
    return False


def all_empty(*values):
    for i in values:
        if not is_empty(i):
            return False
    return True


@lru_cache(maxsize=None)
def read_excel_file(path, sheet_name, row_num_for_column, nrows, with_merged_cells) -> pd.DataFrame:
    """
    :param path:
    :param sheet_name:
    :param row_num_for_column:
    :param nrows:
    :param with_merged_cells:
    :return: [{
        "path": "",
        "sheet_name": "",
        "row_num_for_column": 0,
        "nrows": 1,
        "with_merged_cells": False
    }]
    """
    # if str((path, sheet_name, row_num_for_column, nrows)) in d:
    #     return d[str(("df", path, sheet_name, row_num_for_column, nrows))]
    # path = file_config.get("path")
    # sheet_name = file_config.get("sheet_name")
    # header = int(file_config.get("row_num_for_column")) - 1
    # nrows = file_config.get("nrows")
    header = None
    if row_num_for_column is not None:
        header = int(row_num_for_column) - 1
    df = pd.read_excel(path, sheet_name=sheet_name, header=header, nrows=nrows)
    if with_merged_cells:
        wb = load_workbook(filename=path)
        # 选择要操作的sheet
        sheet = wb[sheet_name]
        # 获取所有合并单元格的信息
        merged_cells = sheet.merged_cells.ranges
        df.merged_cells = merged_cells
    # d[str(("df", path, sheet_name, row_num_for_column, nrows))] = df
    return df


@lru_cache(maxsize=None)
def read_excel_sheets(path, *args, **kwargs) -> typing.List[str]:
    # path = file_config.get("path")
    # if str(("sheets", path, sheet_name, row_num_for_column, nrows)) in d:
    #     return d[str(("sheets", path, sheet_name, row_num_for_column, nrows))]
    sheet_names = pd.ExcelFile(path).sheet_names
    # d[str(("sheets", path, sheet_name, row_num_for_column, nrows))] = sheet_names
    return sheet_names


@lru_cache(maxsize=None)
def read_excel_columns(path, sheet_name, row_num_for_column, *args, **kwargs) -> typing.List[str]:
    # path = file_config.get("path")
    # sheet_name = file_config.get("sheet_name")
    # row_num_for_column = file_config.get("row_num_for_column")
    # if str(("columns", path, sheet_name, row_num_for_column, nrows)) in d:
    #     return d[str(("columns", path, sheet_name, row_num_for_column, nrows))]
    if path.endswith(".xlsx"):
        # 加载工作簿
        wb = load_workbook(filename=path, read_only=True)
        # 获取第一个工作表
        ws = wb[sheet_name]
        # 读取特定的行
        row_data = [ws.cell(int(row_num_for_column), col).value for col in range(1, 100) if ws.cell(1, col).value]
    else:
        # 打开工作簿
        wb = xlrd.open_workbook(path)
        # 获取第一个工作表
        sheet = wb.sheet_by_name(sheet_name)
        # 读取特定的行
        row_data = sheet.row_values(int(row_num_for_column) - 1)
    # d[str(("columns", path, sheet_name, row_num_for_column, nrows))] = row_data
    return [str(i) for i in row_data]


def read_excel_file_with_multiprocessing(file_configs, only_sheet_name=False, only_column_name=False):
    """
    :param file_configs:
        [{
            "path": "",
            "sheet_name": "",
            "row_num_for_column": 0,  # 标题行
            "nrows": 1,
            "with_merged_cells": False,
        }]
    :param only_sheet_name:
    :param only_column_name:
    :return:
    """
    func = read_excel_file
    if only_sheet_name:
        func = read_excel_sheets
    elif only_column_name:
        func = read_excel_columns

    if len(file_configs) == 1:
        config = file_configs[0]
        return [func(config.get("path"), config.get("sheet_name"), config.get("row_num_for_column") or 1, config.get("nrows"), config.get("with_merged_cells"))]

    file_configs_list = [
        (config.get("path"), config.get("sheet_name"), config.get("row_num_for_column") or 1, config.get("nrows"), config.get("with_merged_cells")) for config
        in file_configs]

    # 多于1个用多进程
    num_cores = cpu_count()  # 获取CPU的核心数
    with Pool(processes=min(num_cores, len(file_configs))) as pool:  # 创建一个包含4个进程的进程池
        results = pool.starmap(func, file_configs_list)  # 将函数和参数列表传递给进程池
    return results
