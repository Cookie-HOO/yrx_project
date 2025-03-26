from yrx_project.scene.process_docs.base import Command, ActionContext


# class SelectCurrentScopeCommand(Command):
#     def office_word_run(self, context: ActionContext):
#         selection = context.selection
#         from win32com.client import constants
#
#         if self.content == "单元格":
#             if selection.Information(constants.wdWithInTable):
#                 cell = selection.Cells(1)
#                 context.selected_range = cell.Range
#         elif self.content == "段落":
#             paragraph = selection.Paragraphs(1)
#             context.selected_range = paragraph.Range

class SelectCurrentScopeCommand(Command):

    def office_word_run(self, context: ActionContext):
        SCOPE_MAP = context.const.get("SCOPE_MAP")
        if self.content in SCOPE_MAP:
            context.selection.Expand(Unit=SCOPE_MAP[self.content])


class SelectRangeCommand(Command):
    def __init__(self, boundary, **kwargs):
        super(SelectRangeCommand, self).__init__(**kwargs)
        self.boundary = boundary

    def office_word_run(self, context: ActionContext):
        BOUNDARY_CHECKS = context.const.get("BOUNDARY_CHECKS")
        BOUNDARY_ACTIONS = context.const.get("BOUNDARY_ACTIONS")
        if not self.boundary or not BOUNDARY_CHECKS.get(self.boundary, lambda _: True)(context.selection):
            return

        # 表格相关操作检查
        if "cell" in self.boundary and context.selection.Range.Tables.Count == 0:
            return

        action = BOUNDARY_ACTIONS.get(self.boundary)
        if action:
            unit, move_type = action
            context.selection.HomeKey(Unit=unit, Extend=move_type)


class SelectUntilCommand(Command):
    PRESET_CONDITIONS = {
        "第一个空行": "\r\r"
    }

    def __init__(self, scope, until, direction, **kwargs):
        super(SelectUntilCommand, self).__init__(**kwargs)
        self.scope = scope
        self.until = until
        self.direction = direction

    def office_word_run(self, context: ActionContext):
        # 作用域检查
        if self.scope == "cell" and context.selection.Range.Tables.Count == 0:
            return

        if self.until == "preset":
            self._handle_preset(context)
        else:
            self._handle_custom(context)

    def _handle_preset(self, context):
        condition = self.PRESET_CONDITIONS.get(self.content)
        if condition:
            find = context.selection.Find
            find.Text = condition
            find.Execute()

    def _handle_custom(self, context):
        if self.content:
            find = context.selection.Find
            find.Text = self.content
            find.Forward = self.direction == "right"
            find.Execute()

