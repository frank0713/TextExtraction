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
    direction = "C:\\temp_zip"
    if not os.path.exists(direction):
        os.makedirs(direction)


def uncompress(file_path, user_extensions):
    """
    The core function to uncompress a file, using 7zip. It can only uncompress
    the selected format files.

    :param file_path: A str. An absolute path.
    :param user_extensions:  A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    """

    # ToDo: 更改7z路径，设为安装路径
    # ToDo: 有密码，跳出
    sevenzip_path = "D:\\textgps\\7z.exe"
    for ext in user_extensions:
        subprocess.call([sevenzip_path, "x", file_path, "*", ext, "-y", "-r", "-oC:\\temp_zip"])


def recursive(user_extensions):
    """
    Uncompress the selected format files recursively. Set the loop less than 10
    in-layer compressed.

    :param user_extensions: A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    """

    time = 1
    while time < 10:
        for f in os.listdir("C:\\temp_zip"):
            ext = os.path.splitext(f)[-1]
            f_path = os.path.join("C:\\temp_zip", f)
            if ext not in ['.zip', '.tar', '.rar', '.7z']:
                time += 1
                continue
            if ext in ['.zip', '.tar', '.rar', '.7z']:
                uncompress(file_path=f_path, user_extensions=user_extensions)
                os.remove(f_path)
                time += 1
                continue


def rm_unzip_files():
    """
    Remove the temporal uncompressed files at the end.
    """

    rmtree("C:\\temp_zip")
