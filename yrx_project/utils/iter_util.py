def dedup_list(l: list) -> list:
    new_list = []
    for i in l:
        if i not in new_list:
            new_list.append(i)
    return new_list


def find_repeat_items(l: list) -> list:
    """返回list中重复的元素
    """
    repeat_items = []
    seen = set()
    for item in l:
        if item not in seen:
            seen.add(item)
        else:
            repeat_items.append(item)
    return repeat_items
