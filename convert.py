# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713


from win32com.client import Dispatch
import os


def convert2docx(path):
    """
    Convert .doc/.docm/.rtf to .docx, using win32com module.
    """

    wordapp = Dispatch("Word.Application")
    doc = wordapp.Documents.Open(path)
    # ToDo: 修改路径。另：初始化时创建
    directions = "C:\\temp_convert"
    if os.path.exists(directions) == False:
        os.makedirs(directions)
    convert_file = os.path.join(directions, "temp.docx")
    doc.SaveAs(convert_file, 12)
    doc.Close()
    wordapp.Quit()


def convert2pptx(path):
    """
    Convert .ppt/.pptm to .pptx, using win32com module.
    """

    pptapp = Dispatch("PowerPoint.Application")
    ppt = pptapp.Presentations.Open(path)
    # ToDo: 解决弹窗
    # ToDo: 修改路径。另：初始化时创建
    # ToDo: 修改路径。另：初始化时创建
    directions = "C:\\temp_convert"
    if os.path.exists(directions) == False:
        os.makedirs(directions)
    convert_file = os.path.join(directions, "temp.pptx")
    ppt.SaveAs(convert_file)
    ppt.Close()
    pptapp.Quit()


def rm_converted_files():
    """Remove all the temporal converted files at the end."""

    temp_files = os.listdir("C:\\temp_convert")
    for f in temp_files:
        os.remove(f)
