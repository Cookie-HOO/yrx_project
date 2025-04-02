# 基础命令类与常量定义
import os

from yrx_project.scene.process_docs.base import ActionContext, Command


# 具体命令实现
class SearchTextCommand(Command):
    def __init__(self, pointer_after_search=None, **kwargs):
        self.pointer_after_search = pointer_after_search
        super().__init__(**kwargs)

    def office_word_check(self, context: ActionContext):
        if self.pointer_after_search and self.pointer_after_search not in context.consts["COLLAPSE_MAP"]:
            raise ValueError("参数非法: pointer_after_search")
        if not self.content:
            raise ValueError("参数非法: content")

    def office_word_run(self, context: ActionContext):
        COLLAPSE_MAP = context.consts["COLLAPSE_MAP"]
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

    def office_word_check(self, context: ActionContext):
        if self.direction not in self.DIRECTION_METHODS:
            raise ValueError("参数非法: direction")
        if self.unit not in context.consts["SCOPE_MAP"]:
            raise ValueError("参数非法: unit")

    def office_word_run(self, context: ActionContext):
        UNIT_TYPES = context.consts["SCOPE_MAP"]
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

    def office_word_check(self, context: ActionContext):
        if self.content not in context.consts["BOUNDARY_MAP"]:
            raise ValueError("参数非法: content")
        if self.direction not in ["left", "right"]:
            raise ValueError("参数非法: direction")
        if not isinstance(self.ignore_blank, bool):
            raise ValueError("参数非法: ignore_blank 必须是布尔值")

    def office_word_run(self, context: ActionContext):
        # 获取范围类型和折叠方向
        boundary_action = context.consts["BOUNDARY_MAP"][self.content]
        unit, collapse_direction = boundary_action

        selection = context.selection

        # 扩展范围到指定的 unit
        range = selection.Range
        range.Expand(Unit=unit)

        if self.select:
            # 选择模式
            if self.ignore_blank:
                success, message = self._select_with_non_blank(context, range)
                if not success:
                    return False, message
            else:
                # 设置选区范围
                if self.direction == "left":
                    range.End = selection.Start
                else:
                    range.Start = selection.End
                selection.SetRange(range.Start, range.End)
        else:
            # 移动模式
            if self.ignore_blank:
                success, message = self._move_with_non_blank(context, range, unit)
                if not success:
                    return False, message

            # 折叠光标到指定方向
            selection.Collapse(Direction=collapse_direction)

        return True, None

    def _move_with_non_blank(self, context, range, unit):
        """在指定范围内跳过空白，直到找到第一个非空白字符"""
        range.Expand(Unit=unit)  # 扩展到目标范围
        find = range.Find
        find.ClearFormatting()
        find.Text = "[^\s]"  # 匹配非空白字符
        find.MatchWildcards = True
        find.Forward = (self.direction == "right")
        find.Wrap = context.consts["FIND_WRAP"]["wdFindStop"]

        found = find.Execute()
        if found:
            context.selection.Collapse(
                Direction=context.consts["COLLAPSE_MAP"][self.direction]
            )
            return True, None
        else:
            return False, "未找到非空白内容"

    def _select_with_non_blank(self, context, range):
        """在指定范围内选择直到第一个非空白字符"""
        range.Expand(Unit=context.consts["UNIT_MAP"][self.content[0]])  # 扩展到目标范围
        find = range.Find
        find.ClearFormatting()
        find.Text = "[^\s]"  # 匹配非空白字符
        find.MatchWildcards = True
        find.Forward = (self.direction == "right")
        find.Wrap = context.consts["FIND_WRAP"]["wdFindStop"]

        found = find.Execute()
        if found:
            # 设置选区范围
            if self.direction == "left":
                range.End = find.Parent.End
            else:
                range.Start = find.Parent.Start
            context.selection.SetRange(range.Start, range.End)
            return True, None
        else:
            return False, "未找到非空白内容"
class InsertSpecialCommand(Command):
    def office_word_check(self, context: ActionContext):
        if self.content not in context.consts["SYMBOL_MAP"]:
            raise ValueError("参数非法: content")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type != context.consts["SELECTION_TYPE"]["wdSelectionIP"]:
            return False, "当前处于选中状态，无法执行插入操作"
        symbol = context.consts["SYMBOL_MAP"][self.content]
        context.selection.InsertBreak(symbol)
        return True, None

class InsertTextCommand(Command):
    def office_word_check(self, context: ActionContext):
        if not self.content:
            raise ValueError("参数非法: content")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type != context.consts["SELECTION_TYPE"]["wdSelectionIP"]:
            return False, "当前处于选中状态，无法执行插入操作"
        context.selection.TypeText(self.content)
        return True, None


class SelectCurrentScopeCommand(Command):
    def office_word_check(self, context: ActionContext):
        if self.content not in context.consts["SCOPE_MAP"]:
            raise ValueError("参数非法: content")

    def office_word_run(self, context: ActionContext):
        context.selection.Expand(context.consts["SCOPE_MAP"][self.content])
        # 单独处理表格单元格
        if self.content == "表格单元格":
            if not context.selection.Information(context.consts["SELECTION_INFO"]["wdWithInTable"]):
                return False, "当前光标不在表格单元格内"
            # 获取当前单元格的范围
            cell_range = context.selection.Cells(1).Range
            context.selection.SetRange(cell_range.Start, cell_range.End)
            return True, None
        # 处理其他范围
        context.selection.Expand(context.consts["SCOPE_MAP"][self.content])
        return True, None


class SelectUntilCommand(Command):
    def __init__(self, direction, scope, until_type, **kwargs):
        self.direction = direction
        self.scope = scope
        self.until_type = until_type
        super().__init__(**kwargs)

    def office_word_check(self, context: ActionContext):
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


class DeleteCommand(Command):
    def office_word_run(self, context: ActionContext):
        # 确保当前有选中内容或光标处于有效状态
        if context.selection.Type == context.consts["SELECTION_TYPE"]["cursor"]:
            return False, "当前没有选中内容，无法删除"
        context.selection.Delete()
        return True, None


class ReplaceTextCommand(Command):
    def office_word_check(self, context: ActionContext):
        if not self.content:
            raise ValueError("参数非法: content")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type == context.consts["SELECTION_TYPE"]["cursor"]:
            return False, "当前未选中内容，无法执行替换操作"
        context.selection.Text = self.content
        return True, None

class UpdateFontCommand(Command):
    CHINESE_FONT_SIZE_MAP = {
        "初号": 42.0,
        "小初": 36.0,
        "一号": 26.0,
        "小一": 24.0,
        "二号": 22.0,
        "小二": 18.0,
        "三号": 16.0,
        "小三": 15.0,
        "四号": 14.0,
        "小四": 12.0,
        "五号": 10.5,
        "小五": 9.0,
        "六号": 7.5,
        "小六": 6.5,
        "七号": 5.5,
        "八号": 5.0
    }
    def __init__(self, attribute, **kwargs):
        self.attribute = attribute
        super().__init__(**kwargs)

    def office_word_check(self, context: ActionContext):
        if self.attribute not in ["family", "size"]:
            raise ValueError("参数非法: attribute")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type not in context.consts["SELECTION_STATE"]:
            return False, "需要选中内容才能修改字体"
        font = context.selection.Font
        if self.attribute == "family":
            font.Name = self.content
        else:
            font.Size = float(self.content[:-2]) if self.content.endswith("pt") else self.CHINESE_FONT_SIZE_MAP.get(self.content)
        return True, None

class AdjustFontSizeCommand(Command):
    def __init__(self, step, **kwargs):
        self.step = step
        super().__init__(**kwargs)

    def office_word_check(self, context: ActionContext):
        if not isinstance(self.step, int):
            raise ValueError("参数非法: step")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type not in  context.consts["SELECTION_TYPE"]["text"]:
            return False, "需要选中内容才能调整字号"
        current_size = context.selection.Font.Size
        new_size = current_size + (self.step * 2)
        context.selection.Font.Size = new_size
        return True, None

class UpdateFontColorCommand(Command):
    def __init__(self, color_mode, **kwargs):
        self.color_mode = color_mode
        super().__init__(**kwargs)

    def office_word_check(self, context: ActionContext):
        if self.color_mode not in ["preset", "custom"]:
            raise ValueError("参数非法: color_mode")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type not in context.consts["SELECTION_TYPE"]["text"]:
            return False, "需要选中内容才能修改颜色"
        if self.color_mode == "preset":
            color_map = {"红色": 0xFF0000, "蓝色": 0x0000FF, "绿色": 0x00FF00}
            color_value = color_map.get(self.content, 0x000000)
        else:
            color_value = self.content
        context.selection.Font.Color = color_value
        return True, None

class UpdateParagraphCommand(Command):
    def __init__(self, attribute, line_spacing_type=None, **kwargs):
        self.attribute = attribute
        self.line_spacing_type = line_spacing_type
        super().__init__(**kwargs)

    def office_word_check(self, context: ActionContext):
        if self.attribute not in ["line_spacing", "alignment"]:
            raise ValueError("参数非法: attribute")

    def office_word_run(self, context: ActionContext):
        if context.selection.Type not in context.consts["SELECTION_TYPE"]["text"]:
            return False, "需要选中内容才能修改段落格式"
        paragraph = context.selection.ParagraphFormat
        if self.attribute == "line_spacing":
            if self.line_spacing_type == "times":
                paragraph.LineSpacingRule = context.consts["ROW_SPACING_MAP"]["倍数行距"]
                paragraph.LineSpacing = float(self.content)
            elif self.line_spacing_type == "min_bounds":
                paragraph.LineSpacingRule = context.consts["ROW_SPACING_MAP"]["最小行距"]
                paragraph.LineSpacing = float(self.content)
            else:
                paragraph.LineSpacingRule = context.consts["ROW_SPACING_MAP"]["固定行距"]
                paragraph.LineSpacing = float(self.content)
        else:
            paragraph.Alignment = context.consts["ALIGN_MAP"][self.content]
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
                # range_obj.InsertAfter("\n")  # 插入换行符或其他分隔符
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