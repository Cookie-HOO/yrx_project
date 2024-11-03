import typing

import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtWidgets import QTableWidgetItem, QWidget, QComboBox, QPushButton, QHBoxLayout, QTableWidget, QCheckBox, \
    QLineEdit, QSizePolicy
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook

from yrx_project.utils.color_util import rgb_to_hex

"""
pyqt5的table组件，的set和get 
"""

# 单元格填充的回调方法，返回QColor对象，QColor(255, 255, 255)
CELL_STYLE_FUNC_TYPE = typing.Callable[[pd.DataFrame, int, int], QColor]
# 单元格组件的回调方法，返回QComboBox对象或None，None表示普通文本
CELL_WIDGET_FUNC_TYPE = typing.Callable[[pd.DataFrame, int, int], typing.Union[None, QComboBox]]


class TableWidgetWrapper:
    def __init__(self, table_widget=None, del_rows_button=False, add_rows_button=False):
        self.table_widget = table_widget or QTableWidget()

        new_table = self.__add_buttons(add_rows_button, del_rows_button)
        if new_table is not None:
            self.table_widget = new_table

    def __add_buttons(self, add_rows_button, del_rows_button):
        if not add_rows_button and not del_rows_button:
            return

        # 获取table所在的布局
        layout = self.table_widget.parent().layout()

        # 获取table在布局中的位置
        index = layout.indexOf(self.table_widget)

        # 创建一个水平布局，并添加两个按钮
        hbox = QHBoxLayout()

        # 创建两个按钮
        if add_rows_button:
            add_rows_button = QPushButton('末尾添加一项', self.table_widget.parent())
            add_rows_button.clicked.connect(self.__add_row)
            hbox.addWidget(add_rows_button)
        if del_rows_button:
            del_rows_button = QPushButton('删除选中项', self.table_widget.parent())
            del_rows_button.clicked.connect(self.__delete_row)
            hbox.addWidget(del_rows_button)

        # 创建一个新的table，并复制原来table的所有内容
        new_table = QTableWidget(self.table_widget.rowCount(), self.table_widget.columnCount(), self.table_widget.parent())
        for i in range(self.table_widget.rowCount()):
            for j in range(self.table_widget.columnCount()):
                new_table.setItem(i, j, self.table_widget.item(i, j))

        # 删除原来的table
        layout.removeWidget(self.table_widget)
        self.table_widget.setParent(None)

        # 在原来table的位置添加新的组合
        layout.insertLayout(index, hbox)
        layout.insertWidget(index + 1, new_table)
        return new_table

    def __add_row(self):
        row_count = self.table_widget.rowCount()
        # 在表格末尾添加一行
        self.table_widget.insertRow(row_count)

    def __delete_row(self):
        # 获取选中的行
        selected_rows = self.table_widget.selectionModel().selectedRows()
        # 从后往前删除，防止行数变化影响删除
        for index in sorted(selected_rows, reverse=True):
            self.table_widget.removeRow(index.row())

    def get_cell_value(self, row: int, column: int) -> typing.Optional[str]:
        # 尝试获取QTableWidgetItem（普通文本）
        item = self.table_widget.item(row, column)
        if item is not None:
            return item.text()

        # 尝试获取QWidget（下拉框）
        widget = self.table_widget.cellWidget(row, column)
        if isinstance(widget, QComboBox):
            return widget.currentText()

        elif isinstance(widget, QCheckBox):  # 全局单选框
            return widget.isChecked()

        # 如果既不是QTableWidgetItem也不是QComboBox，返回None
        return None

    def get_cell_value_with_color(self, row: int, column: int) -> (typing.Optional[str], typing.Optional[str]):
        # 尝试获取QTableWidgetItem（普通文本）
        item = self.table_widget.item(row, column)
        if item is not None:
            background_color = item.background().color()
            if background_color.isValid():
                return item.text(), rgb_to_hex((background_color.red(), background_color.green(), background_color.blue()))
            else:
                return item.text(), None

        # 尝试获取QWidget（下拉框）
        widget = self.table_widget.cellWidget(row, column)
        if isinstance(widget, QComboBox):
            return widget.currentText(), None

        elif isinstance(widget, QCheckBox):  # 全局单选框
            return widget.isChecked(), None

        # 如果既不是QTableWidgetItem也不是QComboBox，返回None
        return None, None

    def fill_data_with_color(
            self,
            df: pd.DataFrame,
            cell_style_func: typing.Callable[[pd.DataFrame, int, int], QColor] = None,
            cell_widget_func: typing.Callable[[pd.DataFrame, int, int], QWidget] = None,
    ):
        cols_to_drop = [i for i in df.columns if str(i).startswith('__')]
        # df删除 no_use_cols 列
        fill_df = df.drop(cols_to_drop, axis=1)
        # 将dataframe的数据写入QTableWidget
        self.table_widget.setRowCount(fill_df.shape[0])
        self.table_widget.setColumnCount(fill_df.shape[1])
        self.table_widget.setHorizontalHeaderLabels([str(i) for i in fill_df.columns])
        for i in range(fill_df.shape[0]):
            for j in range(fill_df.shape[1]):
                item = QTableWidgetItem(str(fill_df.iloc[i, j]))
                if cell_widget_func is not None:
                    item = cell_widget_func(fill_df, i, j) or item

                # 普通文本对象
                if isinstance(item, QTableWidgetItem):
                    if cell_style_func:
                        color = cell_style_func(fill_df, i, j)
                        if color:
                            item.setBackground(QBrush(color))  # 设置背景颜色
                    self.table_widget.setItem(i, j, item)
                # 复杂组件对象：下拉选项
                elif isinstance(item, QComboBox):
                    self.table_widget.setCellWidget(i, j, item)
        return self

    def get_data_as_df(self) -> pd.DataFrame:
        headers = []
        for i in range(self.table_widget.columnCount()):
            header = self.table_widget.horizontalHeaderItem(i)
            if header is not None:
                headers.append(header.text())
            else:
                headers.append(f'Column{i}')
        data = []
        for i in range(self.table_widget.rowCount()):
            row_data = []
            for j in range(self.table_widget.columnCount()):
                item = self.get_cell_value(i, j)
                row_data.append(item or '')
            data.append(row_data)
        df = pd.DataFrame(data, columns=headers)
        return df

    def get_data_as_rows_and_color(self) -> (pd.DataFrame, list):
        headers = []
        for i in range(self.table_widget.columnCount()):
            header = self.table_widget.horizontalHeaderItem(i)
            if header is not None:
                headers.append(header.text())
            else:
                headers.append(f'Column{i}')
        data = []
        colors = []
        for i in range(self.table_widget.rowCount()):
            row_data = []
            color_data = []
            for j in range(self.table_widget.columnCount()):
                item, color = self.get_cell_value_with_color(i, j)
                row_data.append(item or '')
                color_data.append(color)
            data.append(row_data)
            colors.append(color_data)
        df = pd.DataFrame(data, columns=headers)
        return df, colors

    def clear(self):
        """全部清空"""
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(0)

    def clear_content(self):
        """清空内容，保留行列"""
        for i in range(self.table_widget.rowCount()):
            for j in range(self.table_widget.columnCount()):
                self.table_widget.setItem(i, j, QTableWidgetItem(""))

    def add_rich_widget_row(self, row):
        """增加复杂的组件行，支持以下类型
        只读文本: readonly_text:
        可改文本：editable_text: 支持绑定 onchange 事件 const onchange = (row_num, col_num, row, after_change_text) => {}
        下拉框：dropdown
        复选框：checkbox:
        单选框：radio:
        全局单选框：global_radio: 纵向全局唯一
        按钮组：button_group: 支持绑定 onclick事件 const onchange = (row_num, col_num, row) => {}

        rows:
            [{
                "type": "readonly_text",
                "value": "123",
            }, {
                "type": "editable_text",
                "value": "123",
                "onchange": (row_num, col_num, row, after_change_text) => {}
            }, {
                "type": "dropdown",
                "values": ["1", "2", "3"],
                “display_values": ["1", "2", "3"]]
                "cur_index": 0
            }, {
                "type": "global_radio",
                "value": True,
            }, {
                "type": "button_group",
                "values": [{
                   "value: "测试",
                   "onclick": (row_num, col_num, row),
                }],
            }
            ]
        :return:
        """
        nex_row_index = self.table_widget.rowCount()
        self.table_widget.setRowCount(nex_row_index+1)
        for col_index, cell in enumerate(row):
            cell_type = cell.get("type")
            if cell_type == "readonly_text":
                item = QTableWidgetItem(str(cell.get("value")))
                item.setFlags(Qt.ItemIsEnabled)
                self.table_widget.setItem(nex_row_index, col_index, item)
            elif cell_type == "editable_text":
                item = QTableWidgetItem(cell.get("value"))
                item.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled)
                self.table_widget.setItem(nex_row_index, col_index, item)
            elif cell_type == "dropdown":  # todo 多选
                combo_box = QComboBox()
                combo_box.addItems(cell.get("values"))
                combo_box.setCurrentIndex(cell.get("cur_index", 0))
                self.table_widget.setCellWidget(nex_row_index, col_index, combo_box)
            elif cell_type == "global_radio":
                check_box = QCheckBox()
                check_box.setChecked(cell.get("value", False))
                check_box.stateChanged.connect(lambda state, nex_row_index=nex_row_index,col_index=col_index: self.__global_radio_action(state, cur_row_index=nex_row_index, cur_col_index=col_index))
                self.table_widget.setCellWidget(nex_row_index, col_index, check_box)
            elif cell_type == "button_group":
                button_layout = QHBoxLayout()
                for button in cell.get("values"):
                    value = button.get("value")
                    onclick = button.get("onclick")
                    button_widget = QPushButton(value)
                    # button_widget.setSizePolicy(QSizePolicy.Mininum, QSizePolicy.Fixed)
                    # button_widget.setMinimumSize(100, 30)
                    this_row_values = self.get_values_by_row_index(nex_row_index)
                    button_widget.clicked.connect(lambda checked, col_index=col_index, onclick=onclick, nex_row_index=nex_row_index, this_row_values=this_row_values: onclick(nex_row_index, col_index, this_row_values))
                    button_layout.addWidget(button_widget)

                button_container = QWidget()
                button_container.setLayout(button_layout)
                self.table_widget.setRowHeight(nex_row_index, 40)
                self.table_widget.setCellWidget(nex_row_index, col_index, button_container)

    def delete_row(self, row_index):
        self.table_widget.removeRow(row_index)

    def __global_radio_action(self, state, **kwargs):
        cur_row_index = kwargs.get("cur_row_index")
        cur_col_index = kwargs.get("cur_col_index")
        if state == Qt.Checked:
            for row in range(self.table_widget.rowCount()):
                if row != cur_row_index and self.table_widget.cellWidget(row, cur_col_index).isChecked():
                    self.table_widget.cellWidget(row, cur_col_index).setChecked(False)

    def get_columns(self):
        labels = []
        for column in range(self.table_widget.columnCount()):
            item = self.table_widget.horizontalHeaderItem(column)
            if item is not None:
                labels.append(item.text())
        return labels

    def get_values_by_row_index(self, row_index) -> dict:
        values = {}
        for col_index, column_name in enumerate(self.get_columns()):
            widget = self.table_widget.cellWidget(row_index, col_index)
            if isinstance(widget, QTableWidgetItem):
                values[column_name] = widget.text()
            elif isinstance(widget, QComboBox):  # 下拉框
                values[column_name] = widget.currentText()
            elif isinstance(widget, QCheckBox):  # 全局单选框
                values[column_name] = widget.isChecked()
            elif isinstance(widget, QWidget):
                values[column_name] = None
        return values

    def save_with_color(self, path):
        # 创建一个 Workbook
        wb = Workbook()
        ws = wb.active

        # 将 DataFrame 转换为 rows
        df, colors = self.get_data_as_rows_and_color()
        rows = dataframe_to_rows(df, index=False, header=True)

        for i, row in enumerate(rows, 1):
            ws.append(row)
            for j, cell in enumerate(ws[i], 1):
                # 获取对应的颜色
                color = None
                if i >= 2:  # 从第二行开始，即不要表头
                    color = colors[i - 2][j - 1]
                if color:
                    color = color.replace("#", "")
                    # 创建对应的 fill
                    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                    # 设置单元格的背景色
                    cell.fill = fill
        # 保存 workbook
        wb.save(path)
