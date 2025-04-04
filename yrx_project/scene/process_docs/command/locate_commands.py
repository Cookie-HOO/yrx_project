from yrx_project.scene.process_docs.base import Command, ActionContext


class SearchTextCommand(Command):
    def __init__(self, pointer_after_search=None, **kwargs):
        self.pointer_after_search = pointer_after_search  # "left" or "right" or None
        super(SearchTextCommand, self).__init__(**kwargs)

    def office_word_check(self):
        if self.pointer_after_search not in self.consts.get("COLLAPSE_MAP"):
            raise ValueError("参数非法: pointer_after_search")
        if not self.content:
            raise ValueError("参数非法: content")

    def office_word_run(self, context: ActionContext) -> (bool, str):
        # 搜索
        COLLAPSE_MAP = self.consts.get("COLLAPSE_MAP")
        selection = context.selection
        find = selection.Find
        find.ClearFormatting()
        find.Text = self.content
        found = find.Execute()

        # 搜索不到直接返回
        if not found:
            return False, "无法定位，搜索不到"

        # 搜到了的情况下
        # 1. 不设置光标，直接选中
        if self.pointer_after_search is None:  # 直接选中
            selection.Select()  # 选择找到的第一个匹配项
            return True, None

        # 2. 说明搜索完还要移动
        selection.Collapse(Direction=COLLAPSE_MAP.get(self.pointer_after_search))
        return True, None


class MoveCursorCommand(Command):
    DIRECTION_METHODS = {
        'up': 'MoveUp',
        'down': 'MoveDown',
        'left': 'MoveLeft',
        'right': 'MoveRight'
    }

    def __init__(self, unit, direction, **kwargs):
        self.direction = direction.lower()
        self.unit = unit
        super(MoveCursorCommand, self).__init__(**kwargs)
        # self.unit_type = self.DIRECTION_MAPPING.get(self.direction)

    def office_word_check(self):
        if self.direction not in self.DIRECTION_METHODS:
            raise ValueError("参数非法: direction")
        if self.unit not in self.consts.get("SCOPE_MAP"):
            raise ValueError("参数非法: unit")

    def office_word_run(self, context: ActionContext) -> (bool, str):
        # 获取正确的Unit类型
        UNIT_TYPES = self.consts.get("SCOPE_MAP")
        unit = UNIT_TYPES.get(self.unit)

        # 获取对应的方法
        method_name = self.DIRECTION_METHODS.get(self.direction)

        # 执行移动操作
        method = getattr(context.selection, method_name)
        method(Unit=unit, Count=int(self.content))
        return True, None


class MoveCursorUntilSpecialCommand(Command):
    pass