from yrx_project.scene.docs_processor.base import Command, ActionContext


class ReplaceTextCommand(Command):
    def __init__(self, content, **kwargs):
        super(ReplaceTextCommand, self).__init__(**kwargs)
        self.new_content = content

    def office_word_run(self, context: ActionContext):
        if context.selected_range:
            context.selected_range.Text = self.new_content
        else:
            raise RuntimeError("No selection available for replacement")