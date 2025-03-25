from win32com.client import constants

from yrx_project.scene.process_docs.base import Command, ActionContext


# class ReplaceTextCommand(Command):
#
#     def office_word_run(self, context: ActionContext):
#         if context.selected_range:
#             context.selected_range.Text = self.content
#         else:
#             raise RuntimeError("No selection available for replacement")


class ReplaceTextCommand(Command):
    def office_word_run(self, context: ActionContext):
        if context.selection.Start != context.selection.End:
            context.selection.Text = self.content
        else:
            print("未选中任何内容，跳过替换")


class UpdateFontCommand(Command):
    SIZE_MAP = {
        "初号": 42,
        "一号": 26,
        "小三": 15,
        "12pt": 12,
        "14pt": 14,
        "16pt": 16
    }

    def __init__(self, attribute, **kwargs):
        super(UpdateFontCommand, self).__init__(**kwargs)
        self.attribute = attribute

    def office_word_run(self, context: ActionContext):
        if not self.attribute:
            return

        font = context.selection.Font
        if self.attribute == "family" and self.content in ["宋体", "黑体", "仿宋", "楷体", "Times New Roman", "Arial"]:
            font.Name = self.content
        elif self.attribute == "size":
            size = self.SIZE_MAP.get(self.content, 12)
            font.Size = size


class AdjustFontSizeCommand(Command):
    def office_word_run(self, context: ActionContext):
        if self.content in ["增大一级", "减小一级"]:
            step = 1 if self.content == "增大一级" else -1
            context.selection.Font.Size += step


class UpdateFontColorCommand(Command):
    COLOR_MAP = {
        "红色": (255, 0, 0),
        "蓝色": (0, 0, 255),
        "绿色": (0, 255, 0)
    }

    def office_word_run(self, context: ActionContext):
        rgb = self.COLOR_MAP.get(self.content)
        if rgb:
            context.selection.Font.Color = rgb[0] + (rgb[1] << 8) + (rgb[2] << 16)


class UpdateParagraphCommand(Command):
    ALIGN_MAP = {
        "左对齐": constants.wdAlignParagraphLeft,
        "居中": constants.wdAlignParagraphCenter,
        "右对齐": constants.wdAlignParagraphRight,
        "两端对齐": constants.wdAlignParagraphJustify
    }

    def __init__(self, attribute, line_spacing_type=None, **kwargs):
        super(UpdateParagraphCommand, self).__init__(**kwargs)
        self.attribute = attribute
        self.line_spacing_type = line_spacing_type

    def office_word_run(self, context: ActionContext):
        # 行距设置
        if self.attribute == "line_spacing":
            self._set_line_spacing(context)

        # 对齐设置
        elif self.attribute == "alignment" and self.content in self.ALIGN_MAP:
            context.selection.ParagraphFormat.Alignment = self.ALIGN_MAP[self.content]

    def _set_line_spacing(self, context):
        line_type = self.line_spacing_type
        value = float(self.content) if self.content.isdigit() else 1.0

        para = context.selection.ParagraphFormat
        if line_type == "times":
            para.LineSpacingRule = constants.wdLineSpaceMultiple
            para.LineSpacing = value
        elif line_type == "min_bounds":
            para.LineSpacingRule = constants.wdLineSpaceAtLeast
            para.LineSpacing = value
        elif line_type == "fix":
            para.LineSpacingRule = constants.wdLineSpaceExactly
            para.LineSpacing = value