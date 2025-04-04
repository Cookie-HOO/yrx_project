import os


class OfficeWordContext:
    def __init__(self):
        self.word = None
        self.doc = None
        self.selection = None

    def into_file(self, file_path):
        doc = self.word.Documents.Open(os.path.abspath(file_path))
        self.selection = self.word.Selection
        self.doc = doc

    def quit_file(self):
        try:
            self.doc.Save()
        except Exception as e:
            pass
        # doc.Close()
        self.doc = None
        self.selection = None

    @property
    def consts(self) -> dict:
        from win32com.client import constants
        return {
            "SYMBOL_MAP": {
                "分页符": constants.wdPageBreak,
                "换行符": constants.wdLineBreak,
                "分节符": constants.wdSectionBreakNextPage,
                # "制表符": constants.wdTab,
            },
            "ALIGN_MAP": {
                "左对齐": constants.wdAlignParagraphLeft,
                "居中": constants.wdAlignParagraphCenter,
                "右对齐": constants.wdAlignParagraphRight,
                "两端对齐": constants.wdAlignParagraphJustify
            },
            "ROW_SPACING_MAP": {
                "倍数行距": constants.wdLineSpaceMultiple,
                "最小行距": constants.wdLineSpaceAtLeast,
                "固定行距": constants.wdLineSpaceExactly,
            },
            "SCOPE_MAP": {
                "行": constants.wdLine,
                "字符": constants.wdCharacter,
                "段落": constants.wdParagraph,
                "表格单元格": constants.wdCell,
                # "页面": constants.wdPage,
                "整个文档": constants.wdStory
            },
            "BOUNDARY_MAP": {
                "当前行开头": (constants.wdLine, constants.wdCollapseStart),
                "当前行结尾": (constants.wdLine, constants.wdCollapseEnd),
                "上一段开头": (constants.wdParagraph, constants.wdCollapseStart),
                "上一段结尾": (constants.wdParagraph, constants.wdCollapseEnd),
                "当前单元格开头": (constants.wdCell, constants.wdCollapseStart),
                "当前单元格结尾": (constants.wdCell, constants.wdCollapseEnd),
                # "当前页面开头": (constants.wdPage, constants.wdCollapseStart),
                # "当前页面结尾": (constants.wdPage, constants.wdCollapseEnd),
                "当前文档开头": (constants.wdStory, constants.wdCollapseStart),
                "当前文档结尾": (constants.wdStory, constants.wdCollapseEnd),
            },
            "SELECTION_STATE": [
                # "wdSelectionIP": constants.wdSelectionIP,  # 光标状态
                constants.wdSelectionNormal,  # 普通选择
                constants.wdSelectionColumn,  # 选中了一列
                constants.wdSelectionRow,  # 选中了一行或多行
                constants.wdSelectionBlock,  # 选中了一块区域
                constants.wdSelectionInlineShape,  # 选中了内置对象：图片等
                constants.wdSelectionShape,  # 选中了独立的形状
            ],
            "SELECTION_TYPE": {
                # 光标状态
                "cursor": constants.wdSelectionIP,
                # 普通选择, 选中了一列, 选中了一行或多行, 选中了一块区域
                "text": [constants.wdSelectionNormal, constants.wdSelectionColumn, constants.wdSelectionRow, constants.wdSelectionBlock,],
                # 选中了内置对象：图片等, 选中了独立的形状
                "obj": [constants.wdSelectionInlineShape, constants.wdSelectionShape,],
            },
            "SELECTION_INFO": {
                "wdWithInTable": 14,  # 判断光标是否位于表格单元格内
                "wdActiveEndAdjustedPageNumber": 1,  # 获取光标所在页面的页码（调整后）
                "wdActiveEndPageNumber": 3,  # 获取光标所在页面的页码（未调整）
                "wdActiveEndSectionNumber": 2,  # 获取光标所在的节号
                "wdAtEndOfRowMarker": 15,  # 判断光标是否位于表格行的末尾标记处
                "wdFirstCharacterColumnNumber": 9,  # 获取光标所在位置的字符列号
                "wdFirstCharacterLineNumber": 10,  # 获取光标所在位置的行号
                "wdHeaderFooterType": 33,  # 获取光标所在的位置是页眉还是页脚
                "wdHorizontalPositionRelativeToPage": 5,  # 获取光标相对于页面的水平位置
                "wdHorizontalPositionRelativeToTextBoundary": 7,  # 获取光标相对于文本边界的水平位置
                "wdInClipboard": 32,  # 判断剪贴板中是否有内容
                "wdInCommentPane": 28,  # 判断光标是否位于批注窗格中
                "wdInEndnote": 27,  # 判断光标是否位于尾注中
                "wdInFootnote": 26,  # 判断光标是否位于脚注中
                "wdInHeaderFooter": 25,  # 判断光标是否位于页眉或页脚中
                "wdInMasterDocument": 30,  # 判断光标是否位于主文档中
                "wdInWordMail": 31,  # 判断光标是否位于 Word 邮件中
                "wdMaximumNumberOfColumns": 18,  # 获取当前选区的最大列数
                "wdMaximumNumberOfRows": 17,  # 获取当前选区的最大行数
                "wdVerticalPositionRelativeToPage": 6,  # 获取光标相对于页面的垂直位置
                "wdVerticalPositionRelativeToTextBoundary": 8  # 获取光标相对于文本边界的垂直位置
            },
            "FIND_WRAP": {
                "wdFindStop": 0,  # 查找到头后停止
                "wdFindContinue": 1,  # 查找到头后，从另一端继续
            },
            "COLLAPSE_MAP": {
                "left": constants.wdCollapseStart,
                "right": constants.wdCollapseEnd,
            }
        }

    def init(self, debug_mode):
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()

        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = debug_mode
        self.word = word

    def cleanup(self):
        import pythoncom
        pythoncom.CoUninitialize()
        try:
            self.quit_file()
            self.word.Quit()
        except Exception as e:
            pass
        self.word = None
        self.doc = None
        self.selection = None
        del self.word, self.doc, self.selection

