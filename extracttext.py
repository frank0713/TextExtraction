# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713

# ToDo: encode & decode
# ToDo: handle exceptions

import os
from chardet import detect
# office
from win32com.client import Dispatch
from win32com.client import DispatchEx
from docx import Document
from pptx import Presentation
import pandas as pd
# office error
from docx.opc import exceptions
from pptx import exc
from xlrd.biffh import XLRDError
# scan pdf
from PIL import Image
import pytesseract
# doc pdf
from pdf2image import convert_from_path
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
# markup
from xml.etree.ElementTree import ElementTree, ParseError
from bs4 import BeautifulSoup
from tex2py import tex2py
# odf
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

    def csvtext(self):
        """Extract text from .csv files.
        
        Use Pandas library to convert into .txt format."""

        try:
            df = pd.read_csv(self.path, header=None)
            # ToDo: 修改创建dir。另：初始化时创建
            directory = "C:\\temp"
            if not os.path.exists(directory):
                os.makedirs(directory)
            df.to_csv("C:\\temp\\temp.txt", header=None, sep=' ', index=False)
            text = self.ctxttext(file="C:\\temp\\temp.txt")
        except (pd.errors.EmptyDataError, UnicodeDecodeError):
            # empty .csv(also: locked file) or decode error
            text = ''
        return text

    def exceltext(self):
        """Extract text from .xls/.xlsx/.xlsm files.
        
        Use Pandas library to convert the MS-Excel associated format files
        (including old version[.xls] or macro[.xlsm] ones) into .txt format."""

        try:
            df = pd.read_excel(self.path, header=None)
            # ToDo: 修改创建路径。另：初始化时创建
            directory = "C:\\temp"
            if not os.path.exists(directory):
                os.makedirs(directory)
            df.to_csv("C:\\temp\\temp.txt", header=None, sep=' ', index=False)
            text = self.ctxttext(file="C:\\temp\\temp.txt")
        except XLRDError:  # empty or locked file
            text = ''
        return text

    @staticmethod
    def rm_txt_files():
        """Remove the temporal .txt files at the end."""

        if os.path.exists("C:\\temp\\temp.txt"):
            os.remove("C:\\temp\\temp.txt")


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

        wordapp = DispatchEx("Word.Application")
        doc = wordapp.Documents.Open(self.path)
        texts = []
        # ToDo: 加密的不报错，弹窗输密码。raise一个来处理?
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

        pptapp = Dispatch("PowerPoint.Application")
        ppt = pptapp.Presentations.Open(self.path, WithWindow=False)
        # ToDo: 加密的不报错，进程不动。raise一个异常？
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
            code = detect(unicode_text)['encoding']
            print(code)
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

    One type is document style, using pdfminer library to parser; the other is 
    scanned type, which is converted to image and extracted by OCR(tesseract).
    """

    def __init__(self, path):
        self.path = path

    def docpdftext(self):
        """Extract text from document type PDF."""

        with open(self.path, 'rb') as fp:
            parser = PDFParser(fp)
            pdf = PDFDocument()
            parser.set_document(pdf)
            pdf.set_parser(parser)
        device = PDFPageAggregator(PDFResourceManager(), laparams=LAParams())
        interpreter = PDFPageInterpreter(PDFResourceManager(), device)
        texts = []
        for p in pdf.get_pages():
            interpreter.process_page(p)
            layout = device.get_result()
            for x in layout:
                if isinstance(x, LTTextBoxHorizontal):
                    t = x.get_text()
                    texts.append(t)
        text = ' '.join(texts)
        text = text.replace("'", "‘")
        return text
    
    def scanpdftext(self):
        """Extract text from scanned PDF by OCR(tesseract)."""

        # ToDo: 修改路径。另：初始化时创建
        directory = "C:\\temp_spdf"
        if not os.path.exists(directory):
            os.makedirs(directory)
        texts = []
        pages = convert_from_path(self.path, 500)
        for p in pages:
            fn = os.path.join(directory, "temp.jpg")
            p.save(fn, "JPEG")
            # ToDo: 添加语言选择设置
            t = pytesseract.image_to_string(Image.open(fn), lang="chi_sim")
            texts.append(t)
        text = ' '.join(texts)
        text = text.replace("'", "‘")
        return text

    @staticmethod
    def rm_spdf():
        """Remove the temporal files of scanned pdf extractor at the end."""

        for f in os.listdir("C:\\temp_spdf"):
            os.remove(f)
