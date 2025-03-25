from PyQt5.QtCore import Qt, QModelIndex
from PyQt5.QtWidgets import QFileSystemModel


class TreeFileWrapper:
    def __init__(self, tree_widget, root_path: str, on_click=None, on_double_click=None):
        self.tree_widget = tree_widget
        self.root_path = root_path
        self.on_click = on_click  # 新增回调函数参数
        self.on_double_click = on_double_click  # 新增回调函数参数
        self.file_model = QFileSystemModel()
        self._build_tree()
        self._connect_signals()  # 新增信号连接

    def _build_tree(self):
        self.file_model.setRootPath(self.root_path)
        self.tree_widget.setModel(self.file_model)
        root_index = self.file_model.index(self.root_path)
        self.tree_widget.setRootIndex(root_index)

        # 启用列隐藏（取消注释并改进）
        for col in range(1, self.file_model.columnCount()):
            self.tree_widget.hideColumn(col)

    def _connect_signals(self):
        """连接树组件的选择信号"""
        # 双击信号
        self.tree_widget.doubleClicked.connect(self._handle_double_click)
        # 单击信号
        self.tree_widget.clicked.connect(self._handle_click)

    def _handle_click(self, index: QModelIndex):
        """处理选择事件"""
        if index.isValid():
            path = self.file_model.filePath(index)
            if self.on_click:
                self.on_click(path)

    def _handle_double_click(self, index: QModelIndex):
        """处理选择事件"""
        if index.isValid():
            path = self.file_model.filePath(index)
            if self.on_double_click:
                self.on_double_click(path)

    def selected(self) -> str:
        """获取当前选中文件的完整路径"""
        indexes = self.tree_widget.selectedIndexes()
        if indexes:
            return self.file_model.filePath(indexes[0])
        return ""
