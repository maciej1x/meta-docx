# -*- coding: utf-8 -*-
"""

@author: Maciej Ulaszewski
mail:ulaszewski.maciej@gmail.com
github:https://github.com/ulaszewskim

"""
import zipfile
from lxml import etree

#==========
#get comments elements from docx file
#==========
def GetCommentElements(docxfile):
    ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    docxZip = zipfile.ZipFile(docxfile)
    commentsXML = docxZip.read('word/comments.xml')
    et = etree.XML(commentsXML)
    comments = et.xpath('//w:comment',namespaces=ns)
    return comments