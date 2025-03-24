from yrx_project.scene.docs_processor.base import Command, ActionContext


class ReplaceTextCommand(Command):

    def office_word_run(self, context: ActionContext):
        if context.selected_range:
            context.selected_range.Text = self.content
        else:
            raise RuntimeError("No selection available for replacement")