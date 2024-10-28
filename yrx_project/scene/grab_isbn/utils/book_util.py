import re

from yrx_project.utils.df_util import is_empty


def has_author_and_in_text(author: str, text: str):
    if not isinstance(author, str) or not isinstance(text, str):
        return False
    author = author.rstrip("译").rstrip("等").rstrip("无")
    if len(author) == 0 or len(text) == 0:
        return False

    # 利用正则进行 ,或者，进行split
    str_list = re.split(r"[,，、]", author)
    for author_one in str_list:
        if author_one not in text and author_one.replace(" ", "") not in text:
            return False
    return True


def use_chine_edition(row):
    publish_edition = row.get("出版\n类别")
    if is_empty(publish_edition) or not isinstance(publish_edition, str):
        raise ValueError("不支持的版本: 出版\n类别 为空")
    publish_edition = publish_edition.strip()

    if publish_edition in ["外文原版教材"]:
        return False
    elif publish_edition in ["中译本", "国内引进版权影印本"]:
        return True
    else:
        raise ValueError(f"不支持的版本: 出版\n类别 为 {publish_edition}")
