import typing

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QMenu, QAction

from yrx_project.utils.iter_util import remove_item_from_list


class CascaderSelectComboBox(QWidget):
    def __init__(self, values, cur_index=None, first_as_none=False):
        super().__init__()
        self.values = values
        self.first_as_none = first_as_none
        self.current_selection_paths = []  # Track all selected paths

        self.button = QPushButton()
        self.menu = QMenu()
        self.actions = []

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.button)
        self.setLayout(self.layout)

        self.button.setMenu(self.menu)

        self.create_menu(self.menu, values, [])

        # Ensure the default selection is checked
        if cur_index is None:
            cur_index = [0]
        self.update_selection(None, cur_index)

    def create_menu(self, menu, values, path):
        for index, value in enumerate(values):
            action = QAction(QIcon(QPixmap(100, 100).fill(Qt.red)), value['label'], self)
            action.setCheckable(True)
            action.triggered.connect(lambda _, this_action=action, this_path=path + [index]: self.update_selection(this_action, this_path))
            menu.addAction(action)
            self.actions.append((action, path + [index]))

            if 'children' in value:
                submenu = QMenu(value['label'], self)
                self.create_menu(submenu, value['children'], path + [index])
                action.setMenu(submenu)

    def update_selection(self, this_action, path):
        # Uncheck
        for action, action_path in self.actions:
            # Uncheck all actions at the current level
            if len(action_path) == len(path) and action_path[:-1] == path[:-1]:
                action.setChecked(False)

            # Uncheck others
            if action_path[0] == path[0]:
                remove_item_from_list(self.current_selection_paths, action_path)
                action.setChecked(False)

        # Check the current action
        if this_action:
            this_action.setChecked(True)

        # Update current selection paths
        if path not in self.current_selection_paths:
            self.current_selection_paths.append(path)

        # Handle first_as_none logic
        if self.first_as_none:
            if path == [0]:  # If the first option is selected
                # Uncheck all other options
                self.current_selection_paths = [path]
                for action, action_path in self.actions:
                    if action_path != [0]:
                        action.setChecked(False)
            else:
                # Uncheck the first option if another option is selected
                self.current_selection_paths = [p for p in self.current_selection_paths if p != [0]]
                for action, action_path in self.actions:
                    if action_path == [0]:
                        action.setChecked(False)

        # Check all parent levels
        for i in range(len(path)):
            for action, action_path in self.actions:
                if action_path == path[:i + 1]:
                    action.setChecked(True)

        # Update button text
        parent_num = len(set([i[0] for i in self.current_selection_paths]))
        if parent_num == 1:
            self.button.setText(
                self.get_label_from_paths([self.current_selection_paths[0]])
            )
        else:
            self.button.setText(f"选中{parent_num}列")

    def get_label_from_paths(self, paths):
        selected_value_in_list = self.get_selected_values_in_list(paths)
        labels = []
        for path_labels in selected_value_in_list:
            labels.append(" > ".join(path_labels))
        return " & ".join(labels)

    def get_selected_values_in_list(self, paths) -> typing.List[typing.List[str]]:
        labels = []
        for path in paths:
            current_values = self.values
            path_labels = []
            for index in path:
                path_labels.append(current_values[index]['label'])
                if 'children' in current_values[index]:
                    current_values = current_values[index]['children']
                else:
                    break
            labels.append(path_labels)
        return labels

    def selected_values(self):
        if self.first_as_none:
            remove_item_from_list(self.current_selection_paths, [0])
        return self.get_selected_values_in_list(self.current_selection_paths)


# Example usage
if __name__ == '__main__':
    import sys

    app = QApplication(sys.argv)
    values = [{'label': '***不从辅助表增加列***'}, {'label': 'zgh', 'children': [{'label': '添加一列'}, {'label': '补充到主表', 'children': [{'label': '开课号'}, {'label': '课程名称'}, {'label': '课程号'}, {'label': '课程类别'}, {'label': '课程性质'}, {'label': '授课方式'}, {'label': '考核方式'}, {'label': '学分'}, {'label': '学时'}, {'label': '开课单位'}, {'label': '单位号'}, {'label': '授课教师'}, {'label': '授课教师工号'}, {'label': '本科学生数'}, {'label': '教材使用情况'}, {'label': '教材类型'}, {'label': '教材名称'}, {'label': '第一作者'}, {'label': '版次'}, {'label': 'ISBN'}, {'label': '出版社'}, {'label': '出版日期'}]}]}, {'label': 'xm', 'children': [{'label': '添加一列'}, {'label': '补充到主表', 'children': [{'label': '开课号'}, {'label': '课程名称'}, {'label': '课程号'}, {'label': '课程类别'}, {'label': '课程性质'}, {'label': '授课方式'}, {'label': '考核方式'}, {'label': '学分'}, {'label': '学时'}, {'label': '开课单位'}, {'label': '单位号'}, {'label': '授课教师'}, {'label': '授课教师工号'}, {'label': '本科学生数'}, {'label': '教材使用情况'}, {'label': '教材类型'}, {'label': '教材名称'}, {'label': '第一作者'}, {'label': '版次'}, {'label': 'ISBN'}, {'label': '出版社'}, {'label': '出版日期'}]}]}, {'label': 'xsh', 'children': [{'label': '添加一列'}, {'label': '补充到主表', 'children': [{'label': '开课号'}, {'label': '课程名称'}, {'label': '课程号'}, {'label': '课程类别'}, {'label': '课程性质'}, {'label': '授课方式'}, {'label': '考核方式'}, {'label': '学分'}, {'label': '学时'}, {'label': '开课单位'}, {'label': '单位号'}, {'label': '授课教师'}, {'label': '授课教师工号'}, {'label': '本科学生数'}, {'label': '教材使用情况'}, {'label': '教材类型'}, {'label': '教材名称'}, {'label': '第一作者'}, {'label': '版次'}, {'label': 'ISBN'}, {'label': '出版社'}, {'label': '出版日期'}]}]}]
    window = CascaderSelectComboBox(values, cur_index=[0], first_as_none=True)
    window.show()
    sys.exit(app.exec_())