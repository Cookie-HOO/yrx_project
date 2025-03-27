from yrx_project.scene.process_docs.base import Command, ActionContext


class InsertSpecialCommand(Command):
    def office_word_run(self, context: ActionContext):
        SYMBOL_MAP = context.office_word_const.get("SYMBOL_MAP")
        if self.content in SYMBOL_MAP:
            context.selection.InsertBreak(SYMBOL_MAP[self.content])


class InsertTextCommand(Command):
    def office_word_run(self, context: ActionContext):
        # 在光标位置插入并保留原选区
        context.selection.TypeText(self.content)

