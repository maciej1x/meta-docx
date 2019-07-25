# -*- coding: utf-8 -*-
"""

@author: Maciej Ulaszewski
mail:ulaszewski.maciej@gmail.com
github:https://github.com/ulaszewskim

"""
import zipfile
from lxml import etree


def get_comment_elements(docxfile):
    """get comments elements from docx file"""
    ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    docx_zip = zipfile.ZipFile(docxfile)
    comments_xml = docx_zip.read('word/comments.xml')
    et = etree.XML(comments_xml)
    comments = et.xpath('//w:comment', namespaces=ns)
    return comments
