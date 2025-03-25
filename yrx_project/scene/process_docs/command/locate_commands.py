from yrx_project.scene.process_docs.base import Command, ActionContext


class SearchTextCommand(Command):
    def __init__(self, pointer_after_search=None, **kwargs):
        super(SearchTextCommand, self).__init__(**kwargs)
        self.pointer_after_search = pointer_after_search  # "left" or "right" or None

    def office_word_run(self, context: ActionContext):
        selection = context.selection
        find = selection.Find
        find.ClearFormatting()
        find.Text = self.content
        found = find.Execute()

        if not found:
            return

        # 搜到了的情况下
        # 1. 不设置光标，直接选中
        if self.pointer_after_search is None:  # 直接选中
            selection.Select()  # 选择找到的第一个匹配项
            return

        # 2. 说明搜索完还要移动
        from win32com.client import constants
        if self.pointer_after_search == "right":
            selection.Collapse(Direction=constants.wdCollapseEnd)
        elif self.pointer_after_search == "left":
            selection.Collapse(Direction=constants.wdCollapseStart)


class MoveCursorCommand(Command):
    DIRECTION_METHODS = {
        'up': 'MoveUp',
        'down': 'MoveDown',
        'left': 'MoveLeft',
        'right': 'MoveRight'
    }

    UNIT_TYPES = {
        'line': 5,  # wdLine
        'character': 1,  # wdCharacter
        'paragraph': 4  # wdParagraph
    }

    # DIRECTION_MAPPING = {
    #     'up': 'line',
    #     'down': 'line',
    #     'left': 'character',
    #     'right': 'character'
    # }

    def __init__(self, unit, direction, **kwargs):
        super(MoveCursorCommand, self).__init__(**kwargs)
        self.direction = direction.lower()
        self.unit = unit
        # self.unit_type = self.DIRECTION_MAPPING.get(self.direction)

    def office_word_run(self, context: ActionContext):
        # 获取正确的Unit类型
        unit = self.UNIT_TYPES.get(self.unit)  # 默认按行

        # 获取对应的方法
        method_name = self.DIRECTION_METHODS.get(self.direction)
        if not method_name:
            raise ValueError(f"Unsupported direction: {self.direction}")

        # 执行移动操作
        method = getattr(context.selection, method_name)
        method(Unit=unit, Count=int(self.content))
