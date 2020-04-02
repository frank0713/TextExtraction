# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713


import os
import subprocess
from shutil import rmtree


# 调用7z.exe


def initial():
    """
    Initialize the directory for storing the temporally uncompressed files.
    """

    # ToDo: 修改路径
    directory = "C:\\temp_zip"
    if not os.path.exists(directory):
        os.makedirs(directory)


def uncompress(file_path, user_extensions):
    """
    The core function to uncompress a file, using 7zip. It can only uncompress
    the selected format files.

    :param file_path: A string. An absolute path.
    :param user_extensions:  A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    """

    # ToDo: 更改7z路径，设为安装路径
    # ToDo: 有密码，跳出
    for suffix in user_extensions:
        ext = "*" + suffix
        # os.system("D:\\textgps\\7z.exe x " + file_path + " " + ext + " -oC:\\temp_zip")
        subprocess.call(["D:\\textgps\\7z.exe", "x", file_path, ext, "-y", "-oC:\\temp_zip"])


def recursive(user_extensions):
    """
    Uncompress the selected format files recursively. 

    :param user_extensions: A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    """

    zip_format = ['.tar', '.rar', '.zip', '.7z']
    intersection = [x for x in zip_format if x in user_extensions]
    for f in os.listdir("C:\\temp_zip"):
        ext = os.path.splitext(f)[-1]
        f_path = os.path.join("C:\\temp_zip", f)
        if ext in intersection:
            uncompress(file_path=f_path, user_extensions=user_extensions)
            os.remove(f_path)


def rm_unzip_files():
    """
    Remove the temporal uncompressed files at the end.
    """

    rmtree("C:\\temp_zip")
