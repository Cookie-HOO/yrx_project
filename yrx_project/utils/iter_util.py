def dedup_list(l: list) -> list:
    new_list = []
    for i in l:
        if i not in new_list:
            new_list.append(i)
    return new_list
