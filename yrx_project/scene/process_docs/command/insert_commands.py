from yrx_project.scene.process_docs.base import Command, ActionContext
from win32com.client import constants


class InsertSpecialCommand(Command):
    SYMBOL_MAP = {
        "分页符": constants.wdPageBreak,
        "换行符": constants.wdLineBreak,
        "分节符": constants.wdSectionBreakNextPage,
        "制表符": constants.wdTab
    }

    def office_word_run(self, context: ActionContext):
        if self.content in self.SYMBOL_MAP:
            context.selection.InsertBreak(self.SYMBOL_MAP[self.content])


class InsertTextCommand(Command):
    def office_word_run(self, context: ActionContext):
        # 在光标位置插入并保留原选区
        context.selection.TypeText(self.content)

