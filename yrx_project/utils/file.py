import os
import shutil
import subprocess
import webbrowser


def copy_file(old_path, new_path, overwrite_if_exists=True):
    if os.path.isdir(old_path) or os.path.isdir(new_path):
        return
    if old_path == new_path:
        return
    if overwrite_if_exists:
        if os.path.exists(new_path):
            print(f"PATH: {new_path} exists, deleting...")
            os.remove(new_path)
    os.makedirs(os.path.dirname(new_path), exist_ok=True)
    shutil.copyfile(old_path, new_path)


def make_zip(source_directory, output_filename):
    """将一个目录下的素有文件和子目录打包"""
    shutil.make_archive(output_filename, 'zip', source_directory)


def get_file_name_without_extension(file_name):
    return os.path.splitext(os.path.basename(file_name))[0]


def get_file_name_with_extension(file_name):
    return os.path.basename(file_name)


def open_file_or_folder_in_browser(file_or_folder_path):
    # file_or_folder_path = '/Users/bytedance/py/yrx_project/yrx_project'  # 替换为你想要打开的文件的路径

    if os.path.exists(file_or_folder_path):
        if os.name == 'nt':  # 对于Windows
            os.startfile(file_or_folder_path)
        elif os.name == 'posix':  # 对于macOS和Linux
            subprocess.call(['open', file_or_folder_path])

    # if os.path.exists(file_or_folder_path):
    #     webbrowser.open(file_or_folder_path)


if __name__ == '__main__':
    open_file_or_folder_in_browser()