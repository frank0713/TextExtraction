# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713

# ToDo: 解编码问题

import os
import convert
from docx import Document
from pptx import Presentation
import pandas as pd
from PIL import Image
import pytesseract
from pdf2image import convert_from_path
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import odf.opendocument as odfopen
from odf import text as odftext
from odf import teletype
from tex2py import tex2py


class TXTText:
    """
    Extract text from general .txt file. 
    
    It can also be used as a txt extractor for some other files, such as: 
    Markup language files(Markdown/Yaml); Code script files(C++/Python...etc.).

    MS-Excel associated files(.xls/.xlsx/.xlsm/.csv) convert into .txt format, 
    using Pandas library, to extract.
    """

    def __init__(self, path):
        self.path = path

    def txttext(self):
        """Extract text from general .txt files."""

        with open(self.path, 'r', encoding='utf-8') as f:
            text = f.read()
            text = text.replace("'", "‘")

        return text

    @staticmethod
    def ctxttext(file):
        """Extract text from converted temporal .txt file for other methods."""

        with open(file, 'r', encoding='utf-8') as f:
            text = f.read()
            text = text.replace("'", "‘")

        return text

    def csvtext(self):
        """Extract text from .csv files.
        
        Use Pandas library to convert into .txt format."""

        df = pd.read_csv(self.path, header=None)
        # ToDo: 修改创建dir。另：初始化时创建
        df.to_csv("C:\\temp\\temp.txt", header=None, sep=' ', index=False)
        text = self.ctxttext(file="C\\temp\\temp.txt")

        return text

    def exceltext(self):
        """Extract text from .xls/.xlsx/.xlsm files.
        
        Use Pandas library to convert the MS-Excel associated format files
        (including old version[.xls] or macro[.xlsm] ones) into .txt format."""

        df = pd.read_excel(self.path, header=None)
        # ToDo: 修改创建路径。另：初始化时创建
        df.to_csv("C:\\temp\\temp.txt", header=None, sep=' ', index=False)
        text = self.ctxttext(file="C:\\temp\\temp.txt")

        return text
    
    @staticmethod
    def rm_files():
        """Remove the temporal files at the end."""

        os.remove("C:\\temp\\temp.txt")


class WordText:
    """
    Extract text from .docx files.

    For old version MS-Word files(.doc), macro-Word files(.docm), or rich text 
    format files(.rtf), they should be converted into .docx format as first, 
    using win32com module (convert2docx).
    """

    def __init__(self, path):
        self.path = path
    
    def docxtext(self):
        """Extract text from .docx files."""

        doc = Document(self.path)
        text = ''
        for p in doc.paragraphs:
            para = p.text
            text = text + para + ' '
        text = text.replace("'", "‘")

        return text

    @staticmethod
    def cdocxtext(file):
        """Extract text from converted temporal .docx file for other methods."""

        doc = Document(file)
        text = ''
        for p in doc.paragraphs:
            para = p.text
            text = text + para + ' '
        text = text.replace("'", "‘")

        return text

    def doctext(self):
        """Extract text from .doc/.docm/.rtf files. They should be converted 
        into .docx at first."""

        convert.convert2docx(self.path)
        text = self.cdocxtext("C\\temp\\temp.docx")

        return text


class PPTText:
    """
    Extract text from .pptx files.

    For old version MS-PowerPoint files(.ppt) or macro-PowerPoint files(.pptm), 
    they should be converted into .pptx format as first, using win32com module
    (convert2pptx).
    """

    def __init__(self, path):
        self.path = path
    
    def pptxtext(self):
        """Extract text from .pptx files."""

        shape_ts = []
        ppt = Presentation(self.path)
        for slide in ppt.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'text'):
                    t = shape.text
                    shape_ts.append(t)
        text = ' '.join(shape_ts)
        text = text.replace("'", "‘")

        return text

    @staticmethod
    def cpptxtext(file):
        """Extract text from converted temporal .pptx file for other methods."""

        shape_ts = []
        ppt = Presentation(file)
        for slide in ppt.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'text'):
                    t = shape.text
                    shape_ts.append(t)
        text = ' '.join(shape_ts)
        text = text.replace("'", "‘")

        return text

    def ppttext(self):
        """Extract text from .ppt/.pptm files. They should be converted 
        into .pptx at first."""

        convert.convert2pptx(self.path)
        text = self.cpptxtext("C:\\temp\\temp.pptx")

        return text


class ODFText:
    """
    Extract text from ODF files(.odt/.ods/.odp).
    """

    def __init__(self, path):
        self.path = path
    
    def odftext(self):
        """Extract text from .odt/.ods/.odp files"""

        odf_file = odfopen.load(self.path)
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

        tree = ET.ElementTree(file=self.path)
        root = tree.getroot()
        texts = []
        for child in root.iter():
            t = child.text
            texts.append(t)
        text = ' '.join(texts)
        text = text.replace("'", "‘")

        return text

    def htmltext(self):
        """Extract text from .html files."""

        with open(self.path, 'r', encoding='utf-8') as hf:
            html = BeautifulSoup(hf, "html.parser")
            text = html.body.get_text()
            text = text.replace("'", "‘")
        
        return text
    
    def textext(self):
        """Extract text from .tex files."""

        with open(self.path, encoding='utf-8') as tf:
            toc = tex2py(tf.read())
        text = []
        for i in toc.descendants:
            if isinstance(i, str):
                text.append(i)
        text = ' '.join(text)
        text = text.replace("'", "‘")

        return text
    

class PDFText:
    """
    Extract text from .pdf files.

    One type is document style, using pdfminer library to parser; the other is 
    scanned, which is converted to image and extracted by OCR(tesseract). 
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
        directions = "C:\\temp_spdf"
        if os.path.exists(directions) == False:
            os.makedirs(directions)
        texts = []
        pages = convert_from_path(self.path, 500)
        image_counter = 1
        for p in pages:
            fn = "temp.jpg"
            p.save(fn, "JPEG")
            image_counter += 1
            img_path = os.path.join(directions, fn)
            # ToDo: 添加语言选择设置
            t = pytesseract.image_to_string(Image.open(img_path), lang="chi_sim")
            texts.append(t)
        text = ' '.join(texts)
        text = text.replace("'", "‘")

        return text

    @staticmethod
    def rm_spdf():
        """Remove the temporal files of scanned pdf extractor at the end."""

        for f in os.listdir("C:\\temp_spdf"):
            os.remove(f)


# For uncompressed files, use recursive method to scan.
