from win32com.client import constants

from yrx_project.scene.docs_processor.base import Command, ActionContext


class SelectCurrentCommand(Command):
    def __init__(self, current, **kwargs):
        super(SelectCurrentCommand, self).__init__(**kwargs)
        self.current = current  # "cell" or "paragraph"

    def office_word_run(self, context: ActionContext):
        selection = context.selection
        if self.current == "cell" and selection.Information(constants.wdWithInTable):
            cell = selection.Cells(1)
            context.selected_range = cell.Range
        else:
            paragraph = selection.Paragraphs(1)
            context.selected_range = paragraph.Range