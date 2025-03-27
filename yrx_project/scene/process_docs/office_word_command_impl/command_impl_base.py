from yrx_project.scene.process_docs.base import ActionContext


class OfficeWordImplBase:
    def __init__(self):
        from win32com.client import constants
        self.__office_word_consts = {
            "SYMBOL_MAP": {
                "分页符": constants.wdPageBreak,
                "换行符": constants.wdLineBreak,
                "分节符": constants.wdSectionBreakNextPage,
                "制表符": constants.wdTab
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
                "页面": constants.wdPage,
                "整个文档": constants.wdStory
            },
            "BOUNDARY_MAP": {
                "当前行开头": (constants.wdLine, constants.wdCollapseStart),
                "当前行结尾": (constants.wdLine, constants.wdCollapseEnd),
                "上一段开头": (constants.wdParagraph, constants.wdCollapseStart),
                "上一段结尾": (constants.wdParagraph, constants.wdCollapseEnd),
                "当前单元格开头": (constants.wdCell, constants.wdCollapseStart),
                "当前单元格结尾": (constants.wdCell, constants.wdCollapseEnd),
                "当前页面开头": (constants.wdPage, constants.wdCollapseStart),
                "当前页面结尾": (constants.wdPage, constants.wdCollapseEnd),
                "当前文档开头": (constants.wdStory, constants.wdCollapseStart),
                "当前文档结尾": (constants.wdStory, constants.wdCollapseEnd),
            },
            "SELECTION_TYPE": {
                "wdSelectionIP": constants.wdSelectionIP,
                "wdSelectionNormal": constants.wdSelectionNormal
            },
            "COLLAPSE_MAP": {
                "left": constants.wdCollapseStart,
                "right": constants.wdCollapseEnd,
            }
        }

    def office_word_run(self, context: ActionContext) -> (bool, str):
        """返回是否执行成功
        成功：意味着操作成功执行，如搜索到并选中，如果搜索不到，算失败，但是不报错
        失败：执行直接失败，或操作不生效（如不在单元格中，无法移动到当前单元格开头，或搜索不到指定内容）
        """
        raise NotImplementedError

    def office_word_check(self):
        pass
