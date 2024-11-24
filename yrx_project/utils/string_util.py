import string

import pandas as pd


def remove_punctuation_and_spaces(text):
    if pd.isna(text):
        return ""
    text = str(text)

    # 定义中文标点符号
    chinese_punctuation = '，。、；：？！《》（）【】『』「」“”‘’—…'
    # 合并英文和中文标点符号
    all_punctuation = string.punctuation + chinese_punctuation
    # 创建一个翻译表，该表指定要删除的所有标点符号和空格
    translator = str.maketrans('', '', all_punctuation + ' ')
    # 使用翻译表删除标点符号和空格
    cleaned_text = text.translate(translator)
    return cleaned_text


if __name__ == '__main__':
    print(remove_punctuation_and_spaces("你好，我是一    个机器人。   "))