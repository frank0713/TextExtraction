# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713


import os
from shutil import rmtree
import zipfile
import tarfile
import py7zr

# ToDo: 改进：只临时解压所选格式的文档
# ToDo: 改进：调用dll，支持解压所有格式
# ToDo: unrar


def unzip(path):
    """Uncompress the .zip files into a temporal dir."""

    # ToDo: 修改路径。另：可以在app初始化时统一执行makedirs
    directions = "C:\\temp_uncompress"
    if not os.path.exists(directions):
        os.makedirs(directions)
    zip_file = zipfile.ZipFile(path)
    for f in zip_file.namelist():
        ext = os.path.splitext(f)[-1]
        if ext == '.zip':
            unzip(os.path.join(directions, f))
        if ext == '.tar':
            untar(os.path.join(directions, f))
        if ext == '.7z':
            un7z(os.path.join(directions, f))
        else:
            zip_file.extract(f, directions)
    zip_file.close()


def untar(path):
    """Uncompress the .tar files into a temporal dir."""

    directions = "C:\\temp_uncompress"
    if not os.path.exists(directions):
        os.makedirs(directions)
    tar_file = tarfile.open(path)
    for f in tar_file.getnames():
        ext = os.path.splitext(f)[-1]
        if ext == '.tar':
            untar(os.path.join(directions, f))
        if ext == '.zip':
            unzip(os.path.join(directions, f))
        if ext == '.7z':
            un7z(os.path.join(directions, f))
        else:
            tar_file.extract(f, directions)
    tar_file.close()


def un7z(path):
    """Uncompress .7z files into a temporal dir."""

    directions = "C:\\temp_uncompress"
    if not os.path.exists(directions):
        os.makedirs(directions)
    sevenz_file = py7zr.SevenZipFile(path, mode='r')
    non_7z = []
    for f in sevenz_file.getnames():
        ext = os.path.splitext(f)[-1]
        if ext == '.7z':
            un7z(os.path.join(directions, f))
        if ext == '.zip':
            unzip(os.path.join(directions, f))
        if ext == '.tar':
            untar(os.path.join(directions, f))
        else:
            non_7z.append(f)
    sevenz_file.extract(path=directions, targets=non_7z)
    sevenz_file.close()


def rm_uncompressed_files():
    """Remove all the temporal uncompressed files at the end."""

    uncompress_dirs = "C:\\temp_uncompress"
    rmtree(uncompress_dirs)
