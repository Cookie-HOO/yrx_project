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
        for i in dir(constants):
            print(i)
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
                "cursor": constants.wdSelectionNormal,
                # 普通选择, 选中了一列, 选中了一行或多行, 选中了一块区域
                "text": [constants.wdSelectionNormal, constants.wdSelectionColumn, constants.wdSelectionRow, constants.wdSelectionBlock,],
                # 选中了内置对象：图片等, 选中了独立的形状
                "obj": [constants.wdSelectionInlineShape, constants.wdSelectionShape,],
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

    def init(self):
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()

        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = True
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

