from PyQt5.QtWidgets import QMenu, QAction


class ButtonMenuWrapper:
    """
    将一个普通的 button，附一个 menu
    menu_list: [
        {"type": "menu_action", "name": "选项1", "func": lambda: 1, "tip": "提示1"},
        {"type": "menu", "name": "含有子选项", "tip": "子菜单提示", "children": [
            {"type": "menu_action", "name": "选项1", "func": lambda: 1, "tip": "子选项提示"},
            {"type": "menu_action", "name": "选项2", "func": lambda: 1},
        ]},
        {"type": "menu_spliter"},
        {"type": "menu_action", "name": "选项2", "func": lambda: 1},
    ]
    """
    def __init__(self, window_obj, button_widget, menu_list):
        self.window_obj = window_obj
        self.button_widget = button_widget
        self.menu_list = menu_list

        self.original_tooltips = {}  # 存储原始提示信息（禁用时会覆盖，恢复时需要恢复提示）

        # 初始化并绑定菜单
        self._build_menu()

    def _build_menu(self):
        """
        根据 menu_list 动态生成菜单并绑定到按钮
        """
        main_menu = QMenu(self.window_obj)
        main_menu.setToolTipsVisible(True)  # 启用主菜单的ToolTip显示

        for item in self.menu_list:
            if item["type"] == "menu_action":
                # 普通菜单项：创建 QAction 并设置提示
                action = QAction(item["name"], self.window_obj)
                if "tip" in item:
                    action.setToolTip(item["tip"])  # 设置 tooltip
                action.triggered.connect(item["func"])
                main_menu.addAction(action)

            elif item["type"] == "menu":
                # 子菜单：创建 QMenu 并设置提示（通过 menuAction）
                sub_menu = QMenu(item["name"], self.window_obj)
                sub_menu.setToolTipsVisible(True)  # 启用子菜单的ToolTip显示
                if "tip" in item:
                    sub_menu.menuAction().setToolTip(item["tip"])  # 设置子菜单项的 tooltip
                self._add_submenu_items(sub_menu, item["children"])
                main_menu.addMenu(sub_menu)

            elif item["type"] == "menu_spliter":
                # 分隔线
                main_menu.addSeparator()

        # 将菜单绑定到按钮
        self.button_widget.setMenu(main_menu)

    def _add_submenu_items(self, sub_menu, children):
        """
        递归添加子菜单项
        :param sub_menu: QMenu 对象（已启用ToolTip）
        :param children: 子菜单项列表
        """
        for child in children:
            if child["type"] == "menu_action":
                # 子菜单中的普通项：设置提示
                action = QAction(child["name"], self.window_obj)
                if "tip" in child:
                    action.setToolTip(child["tip"])  # 设置 tooltip
                action.triggered.connect(child["func"])
                sub_menu.addAction(action)

            elif child["type"] == "menu":
                # 嵌套子菜单：设置提示（通过 menuAction）
                nested_menu = QMenu(child["name"], self.window_obj)
                nested_menu.setToolTipsVisible(True)  # 启用嵌套子菜单的ToolTip显示
                if "tip" in child:
                    nested_menu.menuAction().setToolTip(child["tip"])  # 设置嵌套子菜单项的 tooltip
                self._add_submenu_items(nested_menu, child["children"])
                sub_menu.addMenu(nested_menu)

            elif child["type"] == "menu_spliter":
                # 分隔线
                sub_menu.addSeparator()

    def disable_click(self, paths=None, msg=None):
        """
        禁用按钮交互并设置提示
        :param paths: 索引路径列表（例如[[0], [1, 0]]）
        :param msg: 禁用提示信息（如"功能维护中"）
        """
        if paths is None:
            # 禁用整个按钮
            self._store_tooltip(('button',), self.button_widget.toolTip())
            self.button_widget.setEnabled(False)
            if msg:
                self.button_widget.setToolTip(msg)
            self._disable_all_menu_items(self.button_widget.menu(), msg)
        else:
            # 禁用指定路径
            for path in paths:
                action = self._find_action_by_index_path(path)
                if action:
                    key = tuple(path)
                    self._store_tooltip(key, action.toolTip())
                    action.setEnabled(False)
                    if msg:
                        action.setToolTip(msg)

    def _store_tooltip(self, path_key, original_tip):
        """保存原始提示信息"""
        if path_key not in self.original_tooltips:
            self.original_tooltips[path_key] = original_tip

    def _disable_all_menu_items(self, menu, msg=None):
        """递归禁用所有菜单项并设置提示"""
        if not menu:
            return

        for idx, action in enumerate(menu.actions()):
            if action.isSeparator():
                continue

            # 存储原始提示
            path = self._get_action_path(action)
            if path:
                self._store_tooltip(path, action.toolTip())

            # 设置禁用状态和提示
            action.setEnabled(False)
            if msg:
                action.setToolTip(msg)

            # 递归处理子菜单
            if action.menu():
                self._disable_all_menu_items(action.menu(), msg)

    def enable_click(self):
        """启用所有功能并恢复提示"""
        # 恢复按钮状态
        self.button_widget.setEnabled(True)
        self._restore_tooltip(('button',))

        # 恢复菜单项
        if self.button_widget.menu():
            self._enable_all_menu_items(self.button_widget.menu())

    def _enable_all_menu_items(self, menu):
        """递归启用菜单项并恢复提示"""
        if not menu:
            return

        for action in menu.actions():
            if action.isSeparator():
                continue

            # 恢复提示
            path = self._get_action_path(action)
            if path:
                self._restore_tooltip(path)

            # 启用菜单项
            action.setEnabled(True)

            # 递归处理子菜单
            if action.menu():
                self._enable_all_menu_items(action.menu())

    def _restore_tooltip(self, path_key):
        """恢复原始提示信息"""
        if path_key in self.original_tooltips:
            original = self.original_tooltips.pop(path_key)
            if path_key == ('button',):
                self.button_widget.setToolTip(original)
            else:
                action = self._find_action_by_index_path(list(path_key))
                if action:
                    action.setToolTip(original)

    def _find_action_by_index_path(self, path):
        """根据索引路径查找菜单项"""

        def search_menu(menu, depth=0):
            if not menu or depth >= len(path):
                return None

            current_index = path[depth]
            valid_count = 0

            for action in menu.actions():
                # 跳过分隔线
                if action.isSeparator():
                    continue

                if valid_count == current_index:
                    if depth == len(path) - 1:
                        return action
                    if action.menu():
                        return search_menu(action.menu(), depth + 1)
                    return None
                valid_count += 1

            return None

        return search_menu(self.button_widget.menu())

    def _get_action_path(self, target_action):
        """逆向查找动作的索引路径"""
        path = []
        menu = self.button_widget.menu()
        parent_menu = target_action.parentWidget()

        # 向上追溯菜单层级
        while parent_menu and parent_menu != menu:
            # 在父菜单中查找当前菜单的索引
            parent_parent = parent_menu.parentWidget()
            if not parent_parent:
                break

            index = 0
            for action in parent_parent.actions():
                if action.menu() == parent_menu:
                    path.insert(0, index)
                    break
                if not action.isSeparator():
                    index += 1
            parent_menu = parent_parent

        # 添加顶层索引
        if parent_menu == menu:
            index = 0
            for action in menu.actions():
                if action == target_action or action.menu() == parent_menu:
                    path.insert(0, index)
                    break
                if not action.isSeparator():
                    index += 1

        return tuple(path) if path else None