import typing

import xlwings as xw

SHEET_TYPE = typing.Union[str, int]
SHEETS_TYPE = typing.List[SHEET_TYPE]


class ExcelStyleValue:
    def __init__(self, excel_path, sheet_name_or_index, run_mute=False):
        self.app = xw.App(visible=not run_mute, add_book=False)
        self.wb = self.app.books.open(excel_path)
        self.sht = self.wb.sheets[sheet_name_or_index]

        self.shts = []

    def for_each_sheet(self, func: typing.Callable[[str], None]):
        """遍历存储的sheets，遇到batch_copy时会存储"""
        if not self.shts:
            print("没有存储批量的sheet，请先调用batch copy接口批量创建sheet")
        for sht in self.shts:
            self.sht = sht  # 改变当前self指向
            func(self.sht.name)
        return self

    def get_sheets_name(self):
        return [sht.name for sht in self.wb.sheets]

    # 单元格操作：set、get
    def get_cell(self, cell):
        return self.sht.range(cell).value

    def set_cell(self, cell, value, limit=True):
        if limit:
            self.sht.range(cell).value = value
        return self

    # 单元格操作：合并单元格
    def merge_cell(self, left_top, right_bottom, limit=True):
        """
        :param left_top: (1,2)  表示第一行第二列
        :param right_bottom: 同上
        :param limit: 同上
        :return:
        """
        if limit:
            self.sht.range(left_top, right_bottom).api.MergeCells = True
        return self

    # 行操作：向下复制
    def copy_row_down(self, row_num: int, n=1, set_df=None, limit=True):
        """拷贝这一行往下，的n行，如果给了set_df，就用set_df的值填充拷贝出来的行
        :param row_num:
        :param n:
        :param set_df: 如果行数和拷贝出来的行数一样，只覆盖拷贝出来的行；如果行数比拷贝出来的多一行，那么连最开始的那一行也覆盖
        :param limit: 是否要拷贝
        :return:
        """
        if not limit:
            return self
        for i in range(n):  # 循环粘贴
            row = self.sht.range(f'{row_num}:{row_num}')
            row.api.Copy()  # 拷贝这一行
            self.sht.range(f'{row_num + i+1}:{row_num + i+1}').api.Insert()
        if set_df is not None:
            if n == len(set_df):
                for index, row in set_df.iterrows():
                    self.set_row(row_num + 1 + index, row.to_list())
            elif n == len(set_df) - 1:  # 说明要把最开始的那一行也覆盖掉
                for index, row in set_df.iterrows():
                    self.set_row(row_num + index, row.to_list())
        return self

    def copy_col(self, from_col_num: int, to_col_num, limit=True):
        if not limit:
            return self
        from_col_alpha = self.num2col_char(from_col_num)
        to_col_alpha = self.num2col_char(to_col_num)
        col = self.sht.range(f'{from_col_alpha}:{from_col_alpha}')
        col.api.Copy()
        self.sht.range(f'{to_col_alpha}:{to_col_alpha}').api.Insert()
        return self

    # 行操作：设置一行（保留样式）
    def set_row(self, row_num: int, value_list: typing.List[str]):
        self.sht.range(f'{row_num}:{row_num}').value = value_list
        return self

    def set_col_by_values(self, col_num: int, values: typing.List[str], start_row_num: int = 1):
        """"""
        if len(values) <= 0:
            return self
        start_row_num = start_row_num or 1
        row_nums = range(start_row_num, start_row_num + len(values))
        for row_num, value in zip(row_nums, values):
            self.set_cell((row_num, col_num), value)
        return self

    def set_col_by_col(self, from_col, to_col, start_row_num=1, end_row_num: int = None, end_format: typing.Callable=None):
        start_col_char = self.num2col_char(from_col)
        if end_row_num is None:
            start = f"{start_col_char}1"
            end_row_num = self.sht.range(start).end('down').row
        values = self.sht.range(f"{start_col_char}{start_row_num}:{start_col_char}{end_row_num}").value
        if end_format:
            values = [end_format(i) for i in values]
        self.set_col_by_values(to_col, values, start_row_num=start_row_num)
        return self

    def get_col(self, col_num: int, start_row_num: int, end_row_num: int = None):
        values = []
        if end_row_num is None:
            start = f"{self.num2col_char(col_num)}1"
            end_row_num = self.sht.range().end('down').row

        values = [self.get_cell((start_row_num, col_num))]
        return values

    # 行操作：删除一行（保留样式）
    def delete_row(self, row_num: int, limit=True):
        if limit:
            self.sht.range(f'{row_num}:{row_num}').api.Delete()
        return self

    # 行操作：删除多行（保留样式）
    def batch_delete_row(self, start_row_num: int, end_row_num: int, limit=True):
        if limit:
            self.sht.range(f'{start_row_num}:{end_row_num}').api.Delete()
        return self

    # 列操作：删除一列
    def delete_col(self, col_num: int, limit=True):
        if limit:
            self.sht.range((1, col_num)).api.EntireColumn.Delete()
        return self

    # sheet操作：重命名
    def rename_sheet(self, new_sheet_name: str):
        self.sht.name = new_sheet_name
        return self

    # sheet操作：重命名
    def batch_rename_sheet(self, sheet_name_mapping: typing.Dict[SHEET_TYPE, str]):
        """批量重命名sheet
        :param sheet_name_mapping:
            sheet_name_or_index : new_sheet_name
        :return:
        """
        for sheet_name_or_index, new_sheet_name in sheet_name_mapping.items():
            self.wb.sheets[sheet_name_or_index].name = new_sheet_name
        return self

    # sheet操作：拷贝
    def batch_copy_sheet(self, new_name_list: typing.List[str], append=True, del_old=False):
        """拷贝当前sheet到excel的最后（可以批量拷贝
        :param new_name_list: 可以拷贝多份
        :param append: 如果为True，往后拷贝，结束后，最后一个是activate的；如果是False，从第一个往前拷贝，结束后第一个是activate的
        :param del_old: 如果为True，拷贝完删除老的（剪切）
        :return:
        """
        old_sheet = self.sht
        for name in new_name_list:
            # 复制整个工作表内容
            new_sheet = self.wb.sheets.add(name, after=self.wb.sheets[-1] if append else None)
            old_sheet.api.UsedRange.Copy(new_sheet.api.Range("A1"))

        if del_old:
            old_sheet.delete()

        return self

    def switch_sheet(self, sheet_name_or_index: SHEET_TYPE):
        self.sht = self.wb.sheets[sheet_name_or_index]
        # 这里无法使用activate，需要excel可见
        # self.wb.sheets[sheet_name_or_index].activate()
        return self

    # sheet操作：删除
    def batch_delete_sheet(self, sheet_name_or_index_list: SHEETS_TYPE):
        """将index全都先转成名字，如果有index
        :param sheet_name_or_index_list:
        """
        if len(sheet_name_or_index_list) == 0:
            return self
        sheet_names = [sheet.name for sheet in self.wb.sheets]
        del_sheets = []
        for sheet_name_or_index in sheet_name_or_index_list:
            if isinstance(sheet_name_or_index, int):
                del_sheets.append(sheet_names[sheet_name_or_index])
            else:
                del_sheets.append(sheet_name_or_index)
        for del_sheet_name in set(del_sheets):
            self.wb.sheets[del_sheet_name].delete()
        return self

    def activate_sheet(self, sheet_name_or_index: SHEET_TYPE):
        """保存前可以修改激活的sheet到第一个，打开这个excel就是第一个了"""
        self.wb.sheets[sheet_name_or_index].activate()
        return self

    def save(self, path: str = None):
        # 保存为新的Excel文件
        if path is None:
            self.wb.save()
        else:
            self.wb.save(path)
        self.wb.close()

    def discard(self):
        self.wb.close()

    @staticmethod
    def num2col_char(num: int) -> str:
        """1 -> A"""
        chars = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
        if num > 26:
            raise ValueError("more than 26")
        return chars[num-1]
