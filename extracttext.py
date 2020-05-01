# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713

# ToDo: encode & decode
# ToDo: handle exceptions

import os
# import time
from subprocess import call
from chardet import detect
# **MS Office**
from comtypes.client import CreateObject  # .doc
# from win32com.client import DispatchEx
from docx import Document  # .docx
from pptx import Presentation  # .pptx
from pandas import read_csv, read_excel  # .csv/excel
# **MS Office exceptions**
from _ctypes import COMError  # .doc/.ppt
from docx.opc import exceptions  # .docx
from pptx import exc  # .pptx
from xlrd.biffh import XLRDError  # excel
from pandas.errors import EmptyDataError  # .csv
# **scanned pdf**
# from PIL import Image
from pdf2image import convert_from_path
from pdf2image.exceptions import PDFPageCountError
# import pytesseract
# **doc pdf**
# use xpdf in command line mode
# from pdfplumber import open as pdfopen
# from pdfminer.pdfdocument import PDFPasswordIncorrect
# **Markup formats**
from xml.etree.ElementTree import ElementTree, ParseError
from bs4 import BeautifulSoup
from tex2py import tex2py
# **ODF**
from odf.opendocument import load as odfload
from odf import text as odftext
from odf import teletype


class TXTText:
    """
    Extract text from general .txt file. 
    
    It can also be used as a txt extractor for some other files, such as: 
    Markup language files(Markdown/Yaml); Code script files(C++/Python...etc.).

    MS-Excel associated files(.xls/.xlsx/.xlsm/.csv) convert into .txt format, 
    using Pandas library, to be extracted.
    """

    def __init__(self, path):
        self.path = path

    def txttext(self):
        """Extract text from general .txt files."""

        try:
            with open(self.path, 'rb') as f:
                unicode_text = f.read()
                code = detect(unicode_text)['encoding']  # detect code type
                text = unicode_text.decode(encoding=code)
                text = text.replace("'", "‘")
        except TypeError:  # decode error or empty file(cannot detect code type)
            text = ''
        return text

    @staticmethod
    def ctxttext(file):
        """Extract text from converted temporal .txt file for other inner class
        methods."""

        try:
            with open(file, 'rb') as f:
                unicode_text = f.read()
                code = detect(unicode_text)['encoding']  # detect code type
                print(code)
                text = unicode_text.decode(encoding=code)
                text = text.replace("'", "‘")
        except TypeError:  # decode error or empty file(cannot detect code type)
            text = ''
        return text

    @staticmethod
    def initialize_path():
        """Initialize path for temporal converted .txt files."""

        user_main_path = os.path.expanduser('~')
        directory = user_main_path + '\\Appdata\\Local\\Temp\\Textgps\\temp_txt'
        if not os.path.exists(directory):
            os.makedirs(directory)
        return directory

    def csvtext(self):
        """Extract text from .csv files.
        
        Use Pandas library to convert into .txt format."""

        try:
            df = read_csv(self.path, header=None)
            # ToDo: 修改创建dir。另：初始化时创建
            # 暂使用用户文件夹下路径
            directory = self.initialize_path()
            csv2txt_file = directory + '\\temp.txt'
            df.to_csv(csv2txt_file, header=None, sep=' ', index=False)
            text = self.ctxttext(file=csv2txt_file)
        except (EmptyDataError, UnicodeDecodeError):
            # empty .csv(also: locked file) or decode error
            text = ''
        return text

    def exceltext(self):
        """Extract text from .xls/.xlsx/.xlsm files.
        
        Use Pandas library to convert the MS-Excel associated format files
        (including old version[.xls] or macro[.xlsm] ones) into .txt format."""

        try:
            df = read_excel(self.path, header=None)
            # ToDo: 修改创建路径。另：初始化时创建
            # 暂使用用户文件夹下路径
            directory = self.initialize_path()
            excel2txt_file = directory + '\\.temp.txt'
            df.to_csv(excel2txt_file, header=None, sep=' ', index=False)
            text = self.ctxttext(file=excel2txt_file)
        except XLRDError:  # empty or locked file
            text = ''
        return text

    @staticmethod
    def rm_txt_files():
        """Remove the temporal .txt files at the end."""

        txt_file = TXTText.initialize_path() + 'temp.txt'
        if os.path.exists(txt_file):
            os.remove(txt_file)


class WordText:
    """
    Extract text from .docx/.doc/.docm/.rtf files.

    For .docx files, use docx library.

    For old version MS-Word files(.doc), macro-Word files(.docm), or rich text
    format files(.rtf), use win32com.client module.
    """

    def __init__(self, path):
        self.path = path

    def docxtext(self):
        """Extract text from .docx files."""

        try:
            doc = Document(self.path)
        except exceptions.PackageNotFoundError:
            # cannot open file: locked file, empty file...
            text = ''
        else:
            # main body text
            text = ''
            for p in doc.paragraphs:
                para = p.text
                text = text + para + ' '
            # table text
            table_text = []
            for table in doc.tables:
                for row in range(0, len(table.rows)):
                    r_t = []
                    for column in range(0, len(table.columns)):
                        t = table.cell(row, column).text
                        r_t.append(t)
                    r_t = ' '.join(r_t)
                    table_text.append(r_t)
            table_text = ' \n'.join(table_text)
            text = text + table_text
            text = text.replace("'", "‘")
        return text

    def doctext(self):
        """Extract text from .doc/.docm/.rtf files, using win32com.client."""

        try:
            wordapp = CreateObject("Word.Application")
            doc = wordapp.Documents.Open(self.path, PasswordDocument='')
        except COMError:  # locked file
            text = ''
        else:
            texts = []
            for para in doc.paragraphs:
                t = para.Range.Text
                texts.append(t)
            text = ' '.join(texts)
            text = text.replace("'", "‘")
            doc.Close()
            wordapp.Quit()
        return text


class PPTText:
    """
    Extract text from .pptx/.ppt/.pptm files.

    For .pptx files, use pptx library.

    For old version MS-PowerPoint files(.ppt) or macro-PowerPoint files(.pptm), 
    use win32com.client module.
    """

    def __init__(self, path):
        self.path = path

    def pptxtext(self):
        """Extract text from .pptx files."""

        try:
            shape_ts = []
            ppt = Presentation(self.path)
        except exc.PackageNotFoundError:
            # cannot open file: locked or empty file
            shape_ts = []
        else:
            for slide in ppt.slides:
                for shape in slide.shapes:
                    if hasattr(shape, 'text'):
                        t = shape.text
                        shape_ts.append(t)
        text = ' '.join(shape_ts)
        text = text.replace("'", "‘")
        return text

    def ppttext(self):
        """Extract text from .ppt/.pptm files, using win32com.client module."""

        try:
            pptapp = CreateObject("PowerPoint.Application")
            pwd = ' '
            path_pwd = str(self.path) + "::" + pwd
            ppt = pptapp.Presentations.Open(path_pwd, WithWindow=False)
        except COMError:  # locked file
            text = ''
        else:
            texts = []
            slide_count = ppt.Slides.Count
            for i in range(1, slide_count + 1):
                shape_count = ppt.Slides(i).Shapes.Count
                for j in range(1, shape_count + 1):
                    if ppt.Slides(i).Shapes(j).HasTextFrame:
                        t = ppt.Slides(i).Shapes(j).TextFrame.TextRange.Text
                        texts.append(t)
            text = ' '.join(texts)
            text = text.replace("'", "‘")
            ppt.Close()
            pptapp.Quit()
        return text


class ODFText:
    """
    Extract text from ODF files(.odt/.ods/.odp).
    """

    def __init__(self, path):
        self.path = path

    def odftext(self):
        """Extract text from .odt/.ods/.odp files"""

        odf_file = odfload(self.path)
        odf_text = odf_file.getElementsByType(odftext.P)
        text = ''
        for para in odf_text:
            t = teletype.extractText(para)
            text = text + t + ' '
        text = text.replace("'", "‘")
        return text


class MarkupText:
    """
    Extract text from markup format files(.xml/.html/.tex)

    For .md(markdown) or .yml(yaml) files, use TXT extractor(class TXTText) 
    directly.
    """

    def __init__(self, path):
        self.path = path

    def xmltext(self):
        """Extract text from .xml files."""

        with open(self.path, 'rb') as dxml:
            unicode_text = dxml.read()
            code = detect(unicode_text)['encoding']  # detect code type
            # print(code)
            if code == "None":  # empty file or cannot detect code typy
                text = ''
            else:
                try:
                    with open(file=self.path, encoding=code) as xf:
                        tree = ElementTree(file=xf)
                        root = tree.getroot()
                        texts = []
                        for child in root.iter():
                            t = child.text
                            texts.append(t)
                        text = ' '.join(texts)
                        text = text.replace("'", "‘")
                except ParseError:  # wrong code or others...
                    text = ''
        return text

    def htmltext(self):
        """Extract text from .html files."""

        with open(self.path, 'rb') as dhf:
            unicode_text = dhf.read()
            code = detect(unicode_text)['encoding']
            # print(code)
            if code == "None":  # empty file or connot detect code type
                text = ''
            else:
                try:
                    with open(self.path, 'r', encoding=code) as hf:
                        html = BeautifulSoup(hf, "html.parser")
                        text = html.body.get_text()
                        text = text.replace("'", "‘")
                except AttributeError:  # wrong code or others...
                    text = ''
        return text

    def textext(self):
        """Extract text from .tex files."""

        with open(self.path, 'rb') as dtf:
            unicode_text = dtf.read()
            code = detect(unicode_text)['encoding']
            print(code)
            if code == "None":  # empty file or cannot detect code type
                text = ''
            else:
                try:
                    with open(self.path, 'r', encoding=code) as tf:
                        toc = tex2py(tf.read())
                        text = []
                        for i in toc.descendants:
                            if isinstance(i, str):
                                text.append(i)
                        text = ' '.join(text)
                        text = text.replace("'", "‘")
                except UnicodeDecodeError:  # wrong code type
                    text = ''
        return text


class PDFText:
    """
    Extract text from .pdf files.

    One type is document style, using xpdf to convert to .txt; the other is
    scanned type, which is converted to image and extracted by OCR(tesseract).
    """

    def __init__(self, path):
        self.path = path

    @staticmethod
    def initialize_dpdf_path():
        """Initialize path for temporal doc-pdf files."""

        user_main_path = os.path.expanduser('~')
        dpdf_dir = user_main_path + '\\Appdata\\Local\\Temp\\Textgps\\temp_dpdf'
        if not os.path.exists(dpdf_dir):
            os.makedirs(dpdf_dir)
        return dpdf_dir

    @staticmethod
    def initialize_spdf_path():
        """Initialize path for temporal scan-pdf files."""

        user_main_path = os.path.expanduser('~')
        spdf_dir = user_main_path + '\\Appdata\\Local\\Temp\\Textgps\\temp_spdf'
        if not os.path.exists(spdf_dir):
            os.makedirs(spdf_dir)
        return spdf_dir

    def docpdftext(self):
        """Extract text from document type PDF."""

        # ToDo: 修改xpdf路径
        # ToDo: 修改directory
        xpdf_path = 'D:\\xpdf\\pdftotext.exe'
        directory = self.initialize_dpdf_path()
        file_name = os.path.splitext(os.path.split(self.path)[-1])[0]
        convert_path = directory + file_name + '.txt'
        call([xpdf_path, '-enc', 'UTF-8', self.path, convert_path])
        try:
            with open(convert_path, 'r', encoding='utf-8') as f:
                text = f.read()
                text = text.replace("'", "‘")
        except FileNotFoundError:  # locked file cannot be converted, so it is not found
            text = ''
        return text
    
    def scanpdftext(self):
        """Extract text from scanned PDF by OCR(tesseract)."""

        # ToDo: 修改路径。另：初始化时创建
        directory = self.initialize_spdf_path()
        texts = []
        try:
            pages = convert_from_path(self.path, 300)
            for p in pages:
                fn = os.path.join(directory, "temp.jpg")
                p.save(fn, "JPEG")
                cmd = ['d:\\programs\\tesseract\\tesseract.exe',
                       fn, directory, '--tessdata-dir',
                       'd:\\programs\\tesseract\\tessdata', '-l', 'chi_sim+eng',
                       '--dpi', '300', '--oem', '1']
                call(cmd)
                out_txt = directory + '\\temp.txt'
                with open(out_txt, 'r', encoding='utf-8') as f:
                    t = f.read()
                texts.append(t)
            text = ' '.join(texts)
            text = text.replace("'", "‘")
        except (ValueError, PDFPageCountError):  # cannot parser or others...
            text = ''
        return text

    @staticmethod
    def rm_pdf():
        """Remove the temporal files of scanned pdf extractor at the end."""

        sdpf_path = PDFText.initialize_spdf_path()
        dpdf_path = PDFText.initialize_dpdf_path()
        if os.path.exists(sdpf_path):
            for f in os.listdir(sdpf_path):
                os.remove(f)
        if os.path.exists(dpdf_path):
            for f in os.listdir(sdpf_path):
                os.remove(f)
