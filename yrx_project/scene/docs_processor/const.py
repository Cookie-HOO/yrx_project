ACTION_MAPPING = {
    "定位": {
        "id": "position",
        "children": [
            {"id": "find_first_after", "name": "搜索后光标后移"},
            {"id": "move_down", "name": "光标下移"},
        ],
    },
    "选择": {
        "id": "select",
        "children": [
            {"id": "select_current_cell", "name": "选择当前单元格"},
        ],
    },
    "修改": {
        "id": "update",
        "children": [
            {"id": "replace", "name": "文本替换"},
        ],
    },
    "合并": {
        "id": "n2m",
        "children": [
            {"id": "merge_docs", "name": "合并所有文档"},
        ],
    }
}
