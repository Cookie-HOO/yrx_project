# 基础命令类与常量定义
import os

from yrx_project.scene.process_docs.base import ActionContext, Command


# 具体命令实现
class SearchTextCommand(Command):
    def __init__(self, pointer_after_search=None, **kwargs):
        self.pointer_after_search = pointer_after_search
        super().__init__(**kwargs)

    def office_word_check(self):
        if self.pointer_after_search not in self.consts["COLLAPSE_MAP"]:
            raise ValueError("参数非法: pointer_after_search")
        if not self.content:
            raise ValueError("参数非法: content")

    def office_word_run(self, context: ActionContext):
        COLLAPSE_MAP = self.consts["COLLAPSE_MAP"]
        selection = context.selection
        find = selection.Find
        find.ClearFormatting()
        find.Text = self.content
        found = find.Execute()
        if not found:
            return False, "无法定位，搜索不到"
        if self.pointer_after_search is None:
            selection.Select()
            return True, None
        selection.Collapse(Direction=COLLAPSE_MAP[self.pointer_after_search])
        return True, None

class MoveCursorCommand(Command):
    DIRECTION_METHODS = {
        'up': 'MoveUp',
        'down': 'MoveDown',
        'left': 'MoveLeft',
        'right': 'MoveRight'
    }

    def __init__(self, unit, direction, select=False, **kwargs):
        self.direction = direction.lower()
        self.unit = unit
        self.select = select
        super().__init__(**kwargs)

    def office_word_check(self):
        if self.direction not in self.DIRECTION_METHODS:
            raise ValueError("参数非法: direction")
        if self.unit not in self.consts["SCOPE_MAP"]:
            raise ValueError("参数非法: unit")

    def office_word_run(self, context: ActionContext):
        UNIT_TYPES = self.consts["SCOPE_MAP"]
        method_name = self.DIRECTION_METHODS[self.direction]
        method = getattr(context.selection, method_name)
        method(Count=int(self.content), Unit=UNIT_TYPES[self.unit], Extend=self.select)
        return True, None


class MoveCursorUntilSpecialCommand(Command):
    def __init__(self, direction, ignore_blank, select=False, **kwargs):
        self.direction = direction
        self.ignore_blank = ignore_blank
        self.select = select
        super().__init__(**kwargs)

    def office_word_check(self):
        if self.content not in self.consts["BOUNDARY_ACTIONS"]:
            raise ValueError("参数非法: content")
        if self.direction not in ["left", "right"]:
            raise ValueError("参数非法: direction")
        if not isinstance(self.ignore_blank, bool):
            raise ValueError("参数非法: ignore_blank 必须是布尔值")

    def office_word_run(self, context: ActionContext):
        boundary_action = self.consts["BOUNDARY_ACTIONS"][self.content]
        unit, move_type = boundary_action

        selection = context.selection
        if self.select:
            if self.ignore_blank:
                # 非空白扩展（需结合Find实现，此处简化逻辑）
                selection.Expand(Unit=unit)
            else:
                selection.Expand(Unit=unit)
                selection.Collapse(
                    Direction=self.consts["COLLAPSE_MAP"]["left"] if self.direction == "left" else self.consts["COLLAPSE_MAP"]["right"]
                )
        else:
            if self.ignore_blank:
                # 使用Find跳过空白
                find = selection.Find
                find.ClearFormatting()
                find.Text = ""  # 匹配任意非空白字符
                find.Forward = (self.direction == "right")
                find.Wrap = 0  # wdFindStop
                found = find.Execute()
                if found:
                    selection.Collapse(
                        Direction=self.consts["COLLAPSE_MAP"][self.direction]
                    )
                else:
                    return False, "未找到非空白内容"
            else:
                if self.direction == "left":
                    selection.Collapse(self.consts["COLLAPSE_MAP"]["left"])
                    selection.MoveStart(Unit=unit, Count=1)
                else:
                    selection.Collapse(self.consts["COLLAPSE_MAP"]["right"])
                    selection.MoveEnd(Unit=unit, Count=1)
        return True, None

    def _move_with_non_blank(self, context, unit):
        find = context.selection.Find
        find.ClearFormatting()
        find.Text = ""  # 匹配任意非空白字符
        find.Forward = (self.direction == "right")
        find.Wrap = self.consts["FIND_WRAP"]["wdFindStop"]
        found = find.Execute()
        if found:
            context.selection.Collapse(
                Direction=self.consts["COLLAPSE_MAP"][self.direction]
            )
        else:
            return False, "未找到非空白内容"

class InsertSpecialCommand(Command):
    def office_word_check(self):
        if self.content not in self.consts["SYMBOL_MAP"]:
            raise ValueError("参数非法: content")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type != self.consts["SELECTION_TYPE"]["wdSelectionIP"]:
            return False, "当前处于选中状态，无法执行插入操作"
        symbol = self.consts["SYMBOL_MAP"][self.content]
        context.selection.InsertBreak(symbol)
        return True, None

class InsertTextCommand(Command):
    def office_word_check(self):
        if not self.content:
            raise ValueError("参数非法: content")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type != self.consts["SELECTION_TYPE"]["wdSelectionIP"]:
            return False, "当前处于选中状态，无法执行插入操作"
        context.selection.TypeText(self.content)
        return True, None

class SelectCurrentScopeCommand(Command):
    def __init__(self, scope_type, **kwargs):
        self.scope_type = scope_type
        super().__init__(**kwargs)

    def office_word_check(self):
        if self.scope_type not in self.consts["SCOPE_MAP"]:
            raise ValueError("参数非法: scope_type")

    def office_word_run(self, context: ActionContext):
        context.selection.Expand(self.consts["SCOPE_MAP"][self.scope_type])
        return True, None

class SelectUntilCommand(Command):
    def __init__(self, direction, scope, until_type, **kwargs):
        self.direction = direction
        self.scope = scope
        self.until_type = until_type
        super().__init__(**kwargs)

    def office_word_check(self):
        if self.direction not in ["left", "right"]:
            raise ValueError("参数非法: direction")
        if self.scope not in ["inline", "cell", "paragraph", "page", "doc"]:
            raise ValueError("参数非法: scope")
        if self.until_type not in ["preset", "custom"]:
            raise ValueError("参数非法: until_type")

    def office_word_run(self, context: ActionContext):
        find = context.selection.Find
        find.ClearFormatting()
        if self.until_type == "custom":
            find.Text = self.content
        else:
            find.Text = "\r"
        found = find.Execute()
        if not found:
            return False, "无法找到终止位置"
        context.selection.MoveEndUntil(self.content if self.until_type == "custom" else "\r")
        return True, None

class ReplaceTextCommand(Command):
    def office_word_check(self):
        if not self.content:
            raise ValueError("参数非法: content")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type != self.consts["SELECTION_TYPE"]["wdSelectionNormal"]:
            return False, "当前未选中内容，无法执行替换操作"
        context.selection.Text = self.content
        return True, None

class UpdateFontCommand(Command):
    def __init__(self, attribute, **kwargs):
        self.attribute = attribute
        super().__init__(**kwargs)

    def office_word_check(self):
        if self.attribute not in ["family", "size"]:
            raise ValueError("参数非法: attribute")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type != self.consts["SELECTION_TYPE"]["wdSelectionNormal"]:
            return False, "需要选中内容才能修改字体"
        font = context.selection.Font
        if self.attribute == "family":
            font.Name = self.content
        else:
            font.Size = float(self.content[:-2]) if self.content.endswith("pt") else self.content
        return True, None

class AdjustFontSizeCommand(Command):
    def __init__(self, step, **kwargs):
        self.step = step
        super().__init__(**kwargs)

    def office_word_check(self):
        if not isinstance(self.step, int):
            raise ValueError("参数非法: step")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type != self.consts["SELECTION_TYPE"]["wdSelectionNormal"]:
            return False, "需要选中内容才能调整字号"
        current_size = context.selection.Font.Size
        new_size = current_size + (self.step * 2)
        context.selection.Font.Size = new_size
        return True, None

class UpdateFontColorCommand(Command):
    def __init__(self, color_mode, **kwargs):
        self.color_mode = color_mode
        super().__init__(**kwargs)

    def office_word_check(self):
        if self.color_mode != "preset":
            raise ValueError("参数非法: color_mode")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type != self.consts["SELECTION_TYPE"]["wdSelectionNormal"]:
            return False, "需要选中内容才能修改颜色"
        color_map = {"红色": 0xFF0000, "蓝色": 0x0000FF, "绿色": 0x00FF00}
        context.selection.Font.Color = color_map.get(self.content, 0x000000)
        return True, None

class UpdateParagraphCommand(Command):
    def __init__(self, attribute, line_spacing_type=None, **kwargs):
        self.attribute = attribute
        self.line_spacing_type = line_spacing_type
        super().__init__(**kwargs)

    def office_word_check(self):
        if self.attribute not in ["line_spacing", "alignment"]:
            raise ValueError("参数非法: attribute")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type != self.consts["SELECTION_TYPE"]["wdSelectionNormal"]:
            return False, "需要选中内容才能修改段落格式"
        paragraph = context.selection.ParagraphFormat
        if self.attribute == "line_spacing":
            if self.line_spacing_type == "times":
                paragraph.LineSpacingRule = self.consts["ROW_SPACING_MAP"]["倍数行距"]
                paragraph.LineSpacing = float(self.content)
            elif self.line_spacing_type == "min_bounds":
                paragraph.LineSpacingRule = self.consts["ROW_SPACING_MAP"]["最小行距"]
                paragraph.LineSpacing = float(self.content)
            else:
                paragraph.LineSpacingRule = self.consts["ROW_SPACING_MAP"]["固定行距"]
                paragraph.LineSpacing = float(self.content)
        else:
            paragraph.Alignment = self.consts["ALIGN_MAP"][self.content]
        return True, None


class MergeDocumentsCommand(Command):
    def office_word_run(self, context: ActionContext):
        word = context.word
        new_doc = word.Documents.Add()

        try:
            # 创建一个范围对象，初始指向文档末尾
            range_obj = new_doc.Range(0, 0)

            for file_path in context.input_paths:
                if not os.path.exists(file_path):
                    raise ValueError(f"文件路径不存在: {file_path}")

                # 打开源文件并复制内容
                doc = word.Documents.Open(os.path.abspath(file_path))
                doc.Content.Copy()
                doc.Close(SaveChanges=False)

                # 插入换行符分隔不同文件内容
                range_obj.InsertAfter("\n")  # 插入换行符或其他分隔符
                range_obj.Collapse(Direction=0)  # 移动光标到当前范围的末尾

                # 粘贴内容到指定范围
                range_obj.Paste()

                # 更新范围对象到文档末尾
                range_obj.End = new_doc.Content.End

            # 保存合并后的文档
            output_path = os.path.join(f"{context.command_container.output_folder}",
                                       f"{context.command.action_name}.docx")
            new_doc.SaveAs(os.path.abspath(output_path))
            context.input_paths = [output_path]

        finally:
            new_doc.Close()
            # 确保所有临时文档关闭
            for doc in word.Documents:
                if doc.Name != new_doc.Name:
                    doc.Close(SaveChanges=False)


class MergeDocumentsCommand2(Command):
    def office_word_run(self, context):
        # 创建新文档作为目标
        app = context.selection.Document.Application
        new_doc = app.Documents.Add()

        # 合并当前目录下的所有 .docx 文件（假设路径需调整）
        import glob
        for file_path in glob.glob("*.docx"):
            if file_path != context.selection.Document.FullName:
                new_doc.Content.InsertFile(FileName=file_path)

        # 保存新文档并提示用户
        new_doc.SaveAs("merged_document.docx")  # 可自定义路径
        new_doc.Activate()
        return True, "文档已合并并保存为 merged_document.docx"