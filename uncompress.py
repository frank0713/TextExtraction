# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713


import os
from shutil import rmtree
import zipfile
import tarfile
import py7zr


def unzip(path):
    """Uncompress the .zip files into a temporal dir."""

    # ToDo: 修改路径。另：可以在app初始化时统一执行makedirs
    directions = "C:\\temp_zip"
    if os.path.exists(directions) == False:
        os.makedirs(directions)
    zip_file = zipfile.ZipFile(path)
    for f in zip_file.namelist():
        zip_file.extract(f, directions)
    zip_file.close()


def untar(path):
    """Uncompress the .tar files into a temporal dir."""

    directions = "C:\\temp_tar"
    if os.path.exists(directions) == False:
        os.makedirs(directions)
    tar_file = tarfile.open(path)
    for f in tar_file.getnames():
        tar_file.extract(f, directions)
    tar_file.close()


def un7z(path):
    """Uncompress .7z files into a temporal dir."""

    directions = "C:\\temp_7z"
    if os.path.exists(directions)  == False:
        os.makedirs(directions)
    sevenz_file = py7zr.SevenZipFile(path, mode='r')
    sevenz_file.extractall(directions)
    sevenz_file.close()
    
# ToDo: unrar


def rm_uncompressed_files():
    """Remove all the temporal uncompressed files at the end."""

    unzip_dirs = ["C:\\temp_zip", "C:\\temp_tar", "C:\\temp_7z"]
    for d in unzip_dirs:
        rmtree(d)
