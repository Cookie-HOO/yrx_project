from yrx_project.scene.process_docs.base import Command, ActionContext


class SearchTextCommand(Command):
    def __init__(self, pointer_after_search=None, **kwargs):
        super(SearchTextCommand, self).__init__(**kwargs)
        self.pointer_after_search = pointer_after_search  # "left" or "right" or None

    def office_word_run(self, context: ActionContext):
        COLLAPSE_MAP = context.const.get("COLLAPSE_MAP")
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
        if self.pointer_after_search == "right":
            selection.Collapse(Direction=COLLAPSE_MAP.get("right"))
        elif self.pointer_after_search == "left":
            selection.Collapse(Direction=COLLAPSE_MAP.get("left"))


class MoveCursorCommand(Command):
    DIRECTION_METHODS = {
        'up': 'MoveUp',
        'down': 'MoveDown',
        'left': 'MoveLeft',
        'right': 'MoveRight'
    }


    def __init__(self, unit, direction, **kwargs):
        super(MoveCursorCommand, self).__init__(**kwargs)
        self.direction = direction.lower()
        self.unit = unit
        # self.unit_type = self.DIRECTION_MAPPING.get(self.direction)

    def office_word_run(self, context: ActionContext):
        # 获取正确的Unit类型
        UNIT_TYPES = context.const.get("SCOPE_MAP")
        unit = UNIT_TYPES.get(self.unit)

        # 获取对应的方法
        method_name = self.DIRECTION_METHODS.get(self.direction)
        if not method_name:
            raise ValueError(f"Unsupported direction: {self.direction}")

        # 执行移动操作
        method = getattr(context.selection, method_name)
        method(Unit=unit, Count=int(self.content))


class MoveCursorUntilSpecialCommand(Command):
    pass