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
