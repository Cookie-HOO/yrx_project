import re
import string
import typing

import pandas as pd


# 忽略标点符号
IGNORE_NOTHING = "***不忽略任何内容***"
IGNORE_PUNC = "所有中英文符号和空格"
IGNORE_CHINESE_PAREN = "中文括号及其内容"
IGNORE_ENGLISH_PAREN = "英文括号及其内容"


def remove_chinese_paren(text: str) -> str:
    cleaned_text = re.sub(r'（.*?）', '', text)
    return cleaned_text


def remove_english_paren(text: str) -> str:
    cleaned_text = re.sub(r'\(.*?\)', '', text)
    return cleaned_text


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


def remove_by_ignore_policy(text: str, match_ignore_policy: typing.List[str]) -> str:
    if len(match_ignore_policy) == 0 or IGNORE_NOTHING in match_ignore_policy:
        return text
    if IGNORE_CHINESE_PAREN in match_ignore_policy:
        text = remove_chinese_paren(text)
    if IGNORE_ENGLISH_PAREN in match_ignore_policy:
        text = remove_english_paren(text)
    if IGNORE_PUNC in match_ignore_policy:
        text = remove_punctuation_and_spaces(text)
    return text


if __name__ == '__main__':
    print(remove_punctuation_and_spaces("你好，我是一    个机器人。   "))