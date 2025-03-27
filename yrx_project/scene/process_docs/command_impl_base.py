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
            "BOUNDARY_CHECKS": {
                "cell_start": lambda s: s.Information(constants.wdWithInTable),
                "cell_end": lambda s: s.Information(constants.wdWithInTable),
                "page_start": lambda s: s.Information(constants.wdActiveEndPageNumber) > 0,
            },
            "BOUNDARY_ACTIONS": {
                "line_start": (constants.wdLine, constants.wdMove),
                "line_end": (constants.wdLine, constants.wdExtend),
                "cell_start": (constants.wdCell, constants.wdMove),
                "cell_end": (constants.wdCell, constants.wdExtend),
                "page_start": (constants.wdPage, constants.wdMove),
                "page_end": (constants.wdPage, constants.wdExtend),
                "doc_start": (constants.wdStory, constants.wdMove),
                "doc_end": (constants.wdStory, constants.wdExtend)
            },
            "COLLAPSE_MAP": {
                "right": constants.wdCollapseEnd,
                "left": constants.wdCollapseStart,
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
