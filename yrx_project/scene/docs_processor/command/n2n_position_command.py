from win32com.client import constants

from yrx_project.scene.docs_processor.base import Command, ActionContext


class SearchTextCommand(Command):
    def __init__(self, content, pointer_after_search, **kwargs):
        super(SearchTextCommand, self).__init__(**kwargs)
        self.content = content
        self.pointer_after_search = pointer_after_search  # "before" or "after"


    def office_word_run(self, context: ActionContext):
        selection = context.selection
        find = selection.Find
        find.ClearFormatting()
        find.Text = self.content
        found = find.Execute()

        if found and self.pointer_after_search == "after":
            selection.Collapse(Direction=constants.wdCollapseEnd)
        elif found and self.pointer_after_search == "before":
            selection.Collapse(Direction=constants.wdCollapseStart)
        return found


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

    DIRECTION_MAPPING = {
        'up': 'line',
        'down': 'line',
        'left': 'character',
        'right': 'character'
    }

    def __init__(self, direction, units=1, **kwargs):
        super(MoveCursorCommand, self).__init__(**kwargs)
        self.direction = direction.lower()
        self.units = units
        self.unit_type = self.DIRECTION_MAPPING.get(self.direction)

    def office_word_run(self, context: ActionContext):
        # 获取正确的Unit类型
        unit = self.UNIT_TYPES.get(self.unit_type)  # 默认按行

        # 获取对应的方法
        method_name = self.DIRECTION_METHODS.get(self.direction)
        if not method_name:
            raise ValueError(f"Unsupported direction: {self.direction}")

        # 执行移动操作
        method = getattr(context.selection, method_name)
        method(Unit=unit, Count=self.units)
