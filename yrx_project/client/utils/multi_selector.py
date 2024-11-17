from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QMenu, QAction


class MultiSelectComboBox(QWidget):
    def __init__(self, values, cur_index=0, order=False, bg_colors=None, font_colors=None, first_as_none=False):
        super().__init__()
        self.order = order
        self.values = values
        self.first_as_none = first_as_none

        self.button = QPushButton(values[cur_index])
        self.menu = QMenu()
        self.actions = []
        self.selected_value_list = []

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.button)
        self.setLayout(self.layout)

        self.button.setMenu(self.menu)
        if order:
            self.order_menu = QMenu("调整顺序", self)
            self.order_menu_action = self.menu.addMenu(self.order_menu)
            self.menu.removeAction(self.order_menu_action)  # Hide the order menu initially

        for value in values:
            action = QAction(QIcon(QPixmap(100, 100).fill(Qt.red)), value, self)
            action.setCheckable(True)
            action.triggered.connect(lambda _, this_action=action: self.update_button_text(this_action))
            # self.menu.setStyle()
            self.menu.addAction(action)
            self.actions.append(action)
        self.actions[cur_index].setChecked(True)

    def update_button_text(self, this_action):
        self.selected_value_list = [action.text() for action in self.actions if action.isChecked()]

        if self.order:
            self.button.setText(", ".join(self.selected_value_list))
            if len(self.selected_value_list) > 1:
                if self.order_menu_action not in self.menu.actions():
                    self.menu.addAction(self.order_menu_action)  # Show the order menu
                self.update_order_menu()
            else:
                if self.order_menu_action in self.menu.actions():
                    self.menu.removeAction(self.order_menu_action)  # Hide the order menu
        else:
            if len(self.selected_value_list) == 0:
                if self.first_as_none:  # 如果第一个选项是不操作的选项，则当都不选择时，显示第一个选项
                    self.button.setText(self.values[0])
                    self.actions[0].setChecked(True)  # 为空时，第一项会被选中
                else:
                    self.button.setText(f"选中0项")
            elif len(self.selected_value_list) == 1:
                self.button.setText(self.selected_value_list[0])
            else:
                if self.first_as_none and self.values[0] in self.selected_value_list:
                    if this_action.text() != self.values[0]:  # 如果点击的是普通选项
                        self.actions[0].setChecked(False)
                        self.button.setText(this_action.text())
                    else:
                        self.button.setText(self.values[0])
                        for action in self.actions[1:]:
                            action.setChecked(False)
                else:
                    self.button.setText(f"选中{len(self.selected_value_list)}项")

    def update_order_menu(self):
        self.order_menu.clear()
        for text in self.selected_value_list:
            action = QAction(f"向前移动：{text}", self)
            action.triggered.connect(lambda _, text=text: self.move_item_up(text))
            self.order_menu.addAction(action)

    def move_item_up(self, text):
        items = self.selected_value_list
        index = items.index(text)
        if index > 0:
            items[index], items[index - 1] = items[index - 1], items[index]
            self.button.setText(", ".join(items))

    def selected_values(self):
        if self.first_as_none:  # 如果第一个选项是不操作的选项，那么selected value 里面不包含第一个选项
            return [i for i in self.selected_value_list if i != self.values[0]]
        return self.selected_value_list
