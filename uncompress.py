# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713


import os
from subprocess import call


# use 7z.exe


def initial():
    """
    Initialize the directory for storing the temporally uncompressed files.
    """

    # ToDo: 修改路径
    # 暂使用用户名主路径
    user_main_path = os.path.expanduser('~')
    directory = user_main_path + '\\Appdata\\Local\\Temp\\Textgps\\temp_zip'
    if not os.path.exists(directory):
        os.makedirs(directory)
    return directory


def uncompress(file_path, user_extensions):
    """
    The core function to uncompress a file, using 7zip. It can only uncompress
    the selected formats.

    :param file_path: A string. An absolute path.
    :param user_extensions:  A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    """

    # ToDo: 更改7z路径，设为安装路径
    out_dir = "-o" + initial()
    for suffix in user_extensions:
        ext = "*" + suffix
        # os.system("D:\\textgps\\7z.exe x " + file_path + " " + ext + " -oC:\\temp_zip")
        call(["D:\\textgps\\7z.exe", "x", file_path, ext, "-y", "-p1", out_dir])
        # app_path, extract mode, file_path, extension, all yes, password=1, output_path


def recursive(user_extensions):
    """
    Uncompress the selected format files recursively. 

    :param user_extensions: A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    """

    # Add other formats, if necessary
    zip_format = ['.tar', '.rar', '.zip', '.7z']
    intersection = [x for x in zip_format if x in user_extensions]
    zip_path = initial()
    for f in os.listdir(zip_path):
        ext = os.path.splitext(f)[-1]
        f_path = os.path.join(zip_path, f)
        if ext in intersection:
            uncompress(file_path=f_path, user_extensions=user_extensions)
            os.remove(f_path)


def rm_unzip_files():
    """
    Remove the temporal uncompressed files at the end.
    """

    directory = initial()
    if os.path.exists(directory):
        for f in os.listdir(directory):
            os.remove(f)
