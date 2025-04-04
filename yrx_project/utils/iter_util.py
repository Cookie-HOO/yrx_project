import typing


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


def find_union_and_intersection(l: typing.List[set]) -> (set, set):
    return set.union(*l), set.intersection(*l)


def remove_item_from_list_(l: typing.List[typing.Any], remove: typing.Any) -> typing.List[typing.Any]:
    """
    从l中remove掉第一个remove元素
    l：是一个嵌套list
    """
    try:
        l.remove(remove)
    except ValueError:
        # If 'remove' is not in 'l', do nothing
        pass
    return l


def remove_item_from_list(l: typing.List[typing.Any], remove: typing.Any, iter_delete=False) -> typing.List[typing.Any]:
    """
    从l中remove掉第一个remove元素
    l：是一个嵌套list
    """
    if iter_delete:
        for item in remove:
            remove_item_from_list_(l, item)
    else:
        remove_item_from_list_(l, remove)
    return l


def swap_items_in_origin_list(index1, index2, l) -> list:
    """
    交换l中index1和index2的元素, 原地操作
    注意：负数索引会认为非法，如果出现负数索引，返回原list
    """
    # 参数合法性校验
    length = len(l)
    if length == 0:
        return []
    if index1 < 0 or index2 < 0 or index1 >= length or index2 >= length:
        return l
    if index1 == index2:
        return l
    l[index1], l[index2] = l[index2], l[index1]
    return l

