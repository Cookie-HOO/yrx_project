
class OfficeWordImplBase:
    def office_word_run(self, context: 'ActionContext') -> (bool, str):
        """返回是否执行成功
        成功：意味着操作成功执行，如搜索到并选中，如果搜索不到，算失败，但是不报错
        失败：执行直接失败，或操作不生效（如不在单元格中，无法移动到当前单元格开头，或搜索不到指定内容）
        """
        raise NotImplementedError

    def office_word_check(self, context: 'ActionContext'):
        pass
