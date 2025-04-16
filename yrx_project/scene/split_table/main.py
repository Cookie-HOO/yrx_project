import os
import shutil
import typing

from yrx_project.scene.split_table.const import SCENE_TEMP_PATH
from yrx_project.utils.excel_style import ExcelStyleValue
from yrx_project.utils.file import copy_file


def cleanup_scene_folder():
    if os.path.exists(SCENE_TEMP_PATH):
        shutil.rmtree(SCENE_TEMP_PATH)
    os.makedirs(SCENE_TEMP_PATH)


def split_table(path: str, sheet_name_or_index: typing.Union[str, int], row_num_for_column: int,
                names: typing.List[str], groups: typing.List[typing.Dict[str, str]]) -> str:
    """
    Splits an Excel sheet into multiple sheets based on grouping conditions.

    :param path: Path to the original Excel file.
    :param sheet_name_or_index: Name or index of the sheet to split.
    :param row_num_for_column: Row number containing column headers.
    :param names: List of new sheet names.
    :param groups: List of dictionaries defining grouping conditions.
    :return: Path to the temporary Excel file with split sheets.
    """
    cleanup_scene_folder()
    # Create a temporary copy of the Excel file
    temp_path = os.path.join(SCENE_TEMP_PATH, "split_table.xlsx")
    copy_file(path, temp_path)

    # Initialize Excel interaction
    excel = ExcelStyleValue(temp_path, sheet_name_or_index, run_mute=True)
    original_sheet_name = excel.sht.name

    # Remove all other sheets except the original
    all_sheets = excel.get_sheets_name()
    sheets_to_delete = [s for s in all_sheets if s != original_sheet_name]
    excel.batch_delete_sheet(sheets_to_delete)

    # Validate input lengths
    if len(names) != len(groups):
        raise ValueError("The lengths of names and groups must be the same.")

    # Process each group
    for name, group in zip(names, groups):
        # Copy the original sheet for each new group
        excel.batch_copy_sheet([name], append=True, del_old=False)
        excel.switch_sheet(name)

        # Map headers to column numbers
        header_row = row_num_for_column
        headers = excel.sht.range(f'{header_row}:{header_row}').value
        if not headers or not headers[0]:
            raise ValueError(f"No headers found in row {header_row}.")
        headers = headers[0]
        header_col_map = {header: idx + 1 for idx, header in enumerate(headers)}

        # Prepare conditions
        conditions = {}
        for key, value in group.items():
            if key not in header_col_map:
                raise ValueError(f"Header '{key}' not found in row {header_row}.")
            conditions[header_col_map[key]] = value

        # Determine data rows range
        start_data_row = row_num_for_column + 1
        if conditions:
            first_col = next(iter(conditions.keys()))
            last_row_cell = excel.sht.range((start_data_row, first_col)).end('down')
            last_row = last_row_cell.row
        else:
            last_row = excel.sht.range('A1').end('down').row

        # Filter rows
        current_row = last_row
        while current_row >= start_data_row:
            match = True
            for col_num, expected in conditions.items():
                cell_value = excel.get_cell((current_row, col_num))
                if cell_value != expected:
                    match = False
                    break
            if not match:
                excel.delete_row(current_row)
            current_row -= 1

    # Remove the original sheet
    excel.batch_delete_sheet([original_sheet_name])

    # Save and close
    excel.save()
    excel.discard()

    return temp_path