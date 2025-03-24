from PyQt5.QtWidgets import QMenu, QAction


class ButtonMenuWrapper:
    """
    将一个普通的 button，附一个 menu
    menu_list: [
        {"type": "menu_action", "name": "选项1", "func": lambda: 1},
        {"type": "menu", "name": "含有子选项", "children": [
            {"type": "menu_action", "name": "选项1", "func": lambda: 1},
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

        # 初始化并绑定菜单
        self._build_menu()

    def _build_menu(self):
        """
        根据 menu_list 动态生成菜单并绑定到按钮
        """
        main_menu = QMenu(self.window_obj)

        for item in self.menu_list:
            if item["type"] == "menu_action":
                # 普通菜单项
                action = QAction(item["name"], self.window_obj)
                action.triggered.connect(item["func"])
                main_menu.addAction(action)

            elif item["type"] == "menu":
                # 子菜单
                sub_menu = QMenu(item["name"], self.window_obj)
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
        :param sub_menu: QMenu 对象
        :param children: 子菜单项列表
        """
        for child in children:
            if child["type"] == "menu_action":
                # 普通菜单项
                action = QAction(child["name"], self.window_obj)
                action.triggered.connect(child["func"])
                sub_menu.addAction(action)

            elif child["type"] == "menu":
                # 子菜单
                nested_menu = QMenu(child["name"], self.window_obj)
                self._add_submenu_items(nested_menu, child["children"])
                sub_menu.addMenu(nested_menu)

            elif child["type"] == "menu_spliter":
                # 分隔线
                sub_menu.addSeparator()

