import os

from yrx_project.scene.process_docs.processor import ActionProcessor

test_cases = [
    # 测试搜索相关命令
    {
        "desc": "搜索存在关键词并选中",
        "actions": [
            {'action_id': 'search_first_and_select', 'action_content': "合同条款"},  # 测试SearchTextCommand
        ],
        "expect": "成功选中第一个匹配项"
    },
    {
        "desc": "搜索不存在关键词",
        "actions": [
            {'action_id': 'search_first_and_select', 'action_content': "不存在的文本"},
        ],
        "expect": "返回错误：无法定位，搜索不到"
    },

    # 测试移动光标命令
    {
        "desc": "向上移动2行",
        "actions": [
            {'action_id': 'move_up_lines', 'action_content': 2},  # 测试MoveCursorCommand
        ],
        "expect": "光标上移2行"
    },
    {
        "desc": "无效移动方向参数",
        "actions": [
            {'action_id': 'move_up_lines', 'action_content': "a"},
        ],
        "expect": "参数校验失败：content必须为数字"
    },

    # 测试边界移动命令
    {
        "desc": "向前跳至当前行开头（忽略空白）",
        "actions": [
            {'action_id': 'move_prev_to_landmark_only_text',
             'action_content': "当前行开头"},  # 测试MoveCursorUntilSpecialCommand
        ],
        "expect": "光标移动到当前行第一个非空白字符位置"
    },
    {
        "desc": "向后跳至文档结尾（不忽略空白）",
        "actions": [
            {'action_id': 'move_next_to_landmark',
             'action_content': "当前文档结尾"},
        ],
        "expect": "光标移动到文档最后位置"
    },

    # 测试插入命令
    {
        "desc": "插入分页符",
        "actions": [
            {'action_id': 'insert_special_symbol',
             'action_content': "分页符"},  # 测试InsertSpecialCommand
        ],
        "expect": "文档中插入分页符"
    },
    {
        "desc": "插入自定义文本",
        "actions": [
            {'action_id': 'insert_custom_text',
             'action_content': "测试插入文本"},
        ],
        "expect": "光标位置插入指定文本"
    },

    # 测试选择范围命令
    {
        "desc": "选择当前段落",
        "actions": [
            {'action_id': 'select_current_scope',
             'action_content': "段落"},  # 测试SelectCurrentScopeCommand
        ],
        "expect": "当前段落被选中"
    },
    {
        "desc": "向前选择至表格单元格开头（忽略空白）",
        "actions": [
            {'action_id': 'select_prev_to_landmark_only_text',
             'action_content': "当前单元格开头"},
        ],
        "expect": "选择到单元格第一个非空白字符位置"
    },

    # 测试选择至终止文本
    {
        "desc": "行内选择至指定文本",
        "actions": [
            {'action_id': 'inline_select_to_next_text',
             'action_content': "签名日期"},  # 测试SelectUntilCommand
        ],
        "expect": "选择到下一个出现的'签名日期'位置"
    },
    {
        "desc": "表格内选择至第一个空行",
        "actions": [
            {'action_id': 'cell_select_to_next_condition',
             'action_content': "第一个空行"},
        ],
        "expect": "选择到单元格内第一个空行位置"
    },

    # 测试替换和格式修改
    {
        "desc": "替换选中内容",
        "actions": [
            {'action_id': 'replace_text', 'action_content': "新文本"},  # 测试ReplaceTextCommand
        ],
        "expect": "选中内容被替换为新文本"
    },
    {
        "desc": "设置字体为宋体",
        "actions": [
            {'action_id': 'set_font_family',
             'action_content': "宋体"},  # 测试UpdateFontCommand
        ],
        "expect": "选中文字字体变为宋体"
    },
    {
        "desc": "增大字号一级",
        "actions": [
            {'action_id': 'adjust_font_size',
             'action_content': "增大一级"},  # 测试AdjustFontSizeCommand
        ],
        "expect": "选中文字字号增加2pt"
    },
    {
        "desc": "设置段落居中对齐",
        "actions": [
            {'action_id': 'set_paragraph_alignment',
             'action_content': "居中"},  # 测试UpdateParagraphCommand
        ],
        "expect": "段落对齐方式变为居中"
    },

    # 测试合并文档
    {
        "desc": "合并所有文档",
        "actions": [
            {'action_id': 'merge_documents',
             'action_content': ""},  # 测试MergeDocumentsCommand
        ],
        "expect": "生成合并后的文档文件"
    },

    # 错误测试案例
    {
        "desc": "无效的特殊符号类型",
        "actions": [
            {'action_id': 'insert_special_symbol',
             'action_content': "无效符号"},
        ],
        "expect": "参数校验失败：content非法"
    },
    {
        "desc": "未选中内容时修改格式",
        "actions": [
            {'action_id': 'set_font_color',
             'action_content': "红色"},  # 测试UpdateFontColorCommand
        ],
        "expect": "操作失败：需要选中内容才能修改颜色"
    },
]

ActionProcessor([
    {'action_id': 'search_first_and_select', 'action_content': "职务"},
    {'action_id': 'move_right', 'action_content': 1},
    {'action_id': 'select_current', 'action_content': "单元格"},
    {'action_id': 'replace_text', 'action_content': "sadfasdfsdaf"}
    # {'action_id': 'merge_docs', 'action_params': {'content': '撒扩大飞机阿萨'}}
]).process(file_paths=[
    r"D:\Projects\yrx_project\test1.docx",
    # r"D:\Projects\yrx_project\test2.docx",
    # r"D:\Projects\yrx_project\test3.docx",
])


def integration_test():
    # 测试环境准备
    test_file = "test_template.docx"  # 需要包含以下内容的测试文档：
    # 内容示例：
    # "职务：产品经理\n负责产品需求分析\n联系方式：123456"
    # "表格示例：
    # | 姓名 | 职位 |
    # |------|------|
    # | 张三 | 工程师 |"

    # 预期最终结果：
    # 1. 文档中"产品经理"被替换为"sadfasdfsdaf"
    # 2. 表格单元格内插入分页符
    # 3. 新增合并后的文档包含原始内容和修改内容

    try:
        # 步骤1：搜索"职务"并移动光标到右侧
        ActionProcessor([
            {'action_id': 'search_first_and_move_right', 'action_content': "职务"},
            {'action_id': 'move_right_chars', 'action_content': 2},  # 跳过冒号和空格
        ]).process(file_paths=[test_file])
        # 验证：光标应定位在"产品经理"的起始位置

        # 步骤2：选择当前单元格内容
        ActionProcessor([
            {'action_id': 'select_current_scope', 'action_content': "表格单元格"},
        ]).process(file_paths=[test_file])
        # 验证：当前表格单元格被选中

        # 步骤3：插入分页符
        ActionProcessor([
            {'action_id': 'insert_special_symbol', 'action_content': "分页符"},
        ]).process(file_paths=[test_file])
        # 验证：表格单元格内出现分页符

        # 步骤4：替换文本内容
        ActionProcessor([
            {'action_id': 'search_first_and_select', 'action_content': "产品经理"},
            {'action_id': 'replace_text', 'action_content': "sadfasdfsdaf"}
        ]).process(file_paths=[test_file])
        # 验证："产品经理"被替换为指定文本

        # 步骤5：合并文档
        ActionProcessor([
            {'action_id': 'merge_documents', 'action_content': ""},
        ]).process(file_paths=[test_file])
        # 验证：生成merged_document.docx包含所有修改内容

        # 验证逻辑（伪代码）
        assert check_text_replaced(test_file, "sadfasdfsdaf")  # [[1]][[4]]
        assert check_table_contains(test_file, "分页符")  # [[9]]
        assert os.path.exists("merged_document.docx")  # [[2]][[3]]

        print("集成测试通过")
    except AssertionError as e:
        print(f"测试失败: {str(e)}")
    finally:
        # 清理测试文件
        if os.path.exists("merged_document.docx"):
            os.remove("merged_document.docx")

def check_text_replaced(file_path, target_text):
    # 通过Word API检查是否存在目标文本
    doc = get_word_document(file_path)
    return target_text in doc.Content.Text

def check_table_contains(file_path, symbol):
    doc = get_word_document(file_path)
    for table in doc.Tables:
        for cell in table.Cell:
            if symbol in cell.Range.Text:
                return True
    return False

def get_word_document(file_path):
    from win32com.client import Dispatch
    word = Dispatch('Word.Application')
    return word.Documents.Open(file_path)