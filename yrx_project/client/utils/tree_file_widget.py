import os
from PyQt5.QtCore import Qt, QModelIndex
from PyQt5.QtWidgets import QFileSystemModel, QMenu, QAction, QTreeView, QHeaderView


class TreeFileWrapper:
    def __init__(
            self,
            tree_widget: QTreeView,
            root_path: str,
            on_click=None,
            on_double_click=None,
            right_click_menu=None,
            open_on_default=None
    ):
        """
        :param tree_widget: QTreeView 实例
        :param root_path: 根目录路径
        :param on_click: 单击回调函数
        :param on_double_click: 双击回调函数
        :param right_click_menu: 右键菜单配置
                menu_list: [
                    {"type": "menu_action", "name": "选项1", "func": lambda path: print(f"Clicked {path}")},
                    {"type": "menu_spliter"},
                    {"type": "menu_action", "name": "选项2", "func": lambda path: print(f"Clicked {path}")},
                ]
        :param open_on_default: 默认展开的文件夹设置
                None 表示默认全部关闭
                [] 表示默认全部打开
                或者 list 的 path，表示打开的文件夹
        """
        if not os.path.exists(root_path):
            return
        self.tree_widget = tree_widget
        self.root_path = os.path.normpath(root_path)  # 标准化路径
        self.on_click = on_click
        self.on_double_click = on_double_click
        self.right_click_menu = right_click_menu
        self.open_on_default = open_on_default

        self.file_model = QFileSystemModel()
        self._build_tree()
        self._connect_signals()

        # 启用右键菜单
        if self.right_click_menu:
            self._enable_right_click_menu()

        # 设置默认展开状态
        self.file_model.directoryLoaded.connect(self._set_default_expansion)

    def _build_tree(self):
        """构建树形结构"""
        self.file_model.setRootPath(self.root_path)
        self.tree_widget.setModel(self.file_model)
        root_index = self.file_model.index(self.root_path)
        self.tree_widget.setRootIndex(root_index)

        # # 隐藏多余的列（只保留文件名列）
        # for col in range(1, self.file_model.columnCount()):
        #     self.tree_widget.hideColumn(col)
        # 修改列标题为中文
        # self.tree_widget.header().setStretchLastSection(False)  # 禁用最后一列自动拉伸
        # self.tree_widget.header().setSectionResizeMode(QHeaderView.ResizeToContents)  # 自动调整列宽
        self.tree_widget.header().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)  # 左对齐

        # 设置中文列标题
        self.tree_widget.model().setHeaderData(0, Qt.Horizontal, "名称")
        self.tree_widget.model().setHeaderData(1, Qt.Horizontal, "大小")
        self.tree_widget.model().setHeaderData(2, Qt.Horizontal, "类型")
        self.tree_widget.model().setHeaderData(3, Qt.Horizontal, "修改日期")

        # 调整列宽
        self.tree_widget.setColumnWidth(0, 400)  # 名称列宽度
        self.tree_widget.setColumnWidth(1, 150)  # 大小列宽度
        self.tree_widget.setColumnWidth(2, 200)  # 类型列宽度
        self.tree_widget.setColumnWidth(3, 250)  # 修改日期列宽度

    def _connect_signals(self):
        """连接信号"""
        self.tree_widget.doubleClicked.connect(self._handle_double_click)
        self.tree_widget.clicked.connect(self._handle_click)

    def _handle_click(self, index: QModelIndex):
        """处理单击事件"""
        if index.isValid():
            path = self.file_model.filePath(index)
            if self.on_click:
                self.on_click(path)

    def _handle_double_click(self, index: QModelIndex):
        """处理双击事件"""
        if index.isValid():
            path = self.file_model.filePath(index)
            if self.on_double_click:
                self.on_double_click(path)

    def _enable_right_click_menu(self):
        """启用右键菜单"""
        self.tree_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_widget.customContextMenuRequested.connect(self._show_context_menu)

    def _show_context_menu(self, pos):
        """显示右键菜单"""
        # 获取右键点击位置对应的索引
        index = self.tree_widget.indexAt(pos)
        if not index.isValid():
            return

        # 获取文件路径
        file_path = self.file_model.filePath(index)

        # 创建菜单
        menu = QMenu(self.tree_widget)
        self._build_menu_from_config(menu, self.right_click_menu, file_path)

        # 显示菜单
        menu.exec_(self.tree_widget.viewport().mapToGlobal(pos))

    def _build_menu_from_config(self, menu: QMenu, menu_list, file_path):
        """根据配置动态生成菜单，并将文件路径传递给回调函数"""
        for item in menu_list:
            item_type = item.get("type")
            if item_type == "menu_action":
                action = QAction(item["name"], self.tree_widget)
                # 将文件路径传递给回调函数
                action.triggered.connect(lambda _, path=file_path, func=item["func"]: func(path))
                menu.addAction(action)
            elif item_type == "menu":
                submenu = menu.addMenu(item["name"])
                self._build_menu_from_config(submenu, item["children"], file_path)
            elif item_type == "menu_spliter":
                menu.addSeparator()

    def _set_default_expansion(self):
        """设置默认展开状态"""
        if self.open_on_default is None:
            # 默认全部关闭
            return
        elif isinstance(self.open_on_default, list) and len(self.open_on_default) == 0:
            # 默认全部打开
            self.tree_widget.expandAll()
        else:
            # 展开指定路径的文件夹
            for path in self.open_on_default:
                normalized_path = os.path.normpath(path)  # 标准化路径
                index = self.file_model.index(normalized_path)
                if index.isValid():
                    self.tree_widget.setExpanded(index, True)
                else:
                    print(f"Invalid path: {path}")

    def selected(self) -> str:
        """获取当前选中文件的完整路径"""
        indexes = self.tree_widget.selectedIndexes()
        if indexes:
            return self.file_model.filePath(indexes[0])
        return ""

    def force_refresh(self):
        """
        强制重新检测当前路径。
        调用此方法后，QFileSystemModel 会重新扫描文件系统并更新模型中的数据。
        """
        # 刷新模型
        current_root_path = self.file_model.rootPath()
        self.file_model.setRootPath("")  # 清空根路径以触发刷新
        self.file_model.setRootPath(current_root_path)  # 恢复根路径

        # 更新树形视图
        root_index = self.file_model.index(current_root_path)
        self.tree_widget.setRootIndex(root_index)
