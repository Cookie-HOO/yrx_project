import pandas as pd


def is_empty(value):
    if pd.isna(value) or value is None:
        return True
    return len(str(value)) == 0


def is_not_empty(value):
    return not is_empty(value)


def is_any_empty(*values):
    for i in values:
        if is_empty(i):
            return True
    return False


def all_empty(*values):
    for i in values:
        if not is_empty(i):
            return False
    return True
