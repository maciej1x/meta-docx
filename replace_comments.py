# -*- coding: utf-8 -*-
"""

@author: Maciej Ulaszewski
mail:ulaszewski.maciej@gmail.com
github: https://github.com/ulaszewskim

"""

import os
from shutil import rmtree
from create_new_docx import create_new_docx
from create_temp_zip_dir import create_temp_zip_dir


def replace_comments(coms_to_change, newfile, docxfile):
    """Replace comments in entry docx file and save to other docx file"""
    create_temp_zip_dir(os.path.split(docxfile)[1], os.path.split(docxfile)[0])
    xmlcom = open(os.path.join(os.path.split(docxfile)[0], 'temp_meta_docx/word/comments.xml'), 'r').read()
    for comment in coms_to_change:
        xmlcom = xmlcom.replace(comment[0], comment[1])
    text_file = open(os.path.join(os.path.split(docxfile)[0], 'temp_meta_docx/word/comments.xml'), "w")
    text_file.write(xmlcom)
    text_file.close()
    if os.path.exists(newfile+'.zip'):
        os.remove(newfile+'.zip')
    create_new_docx(os.path.join(os.path.split(docxfile)[0], 'temp_meta_docx/'), newfile)
    rmtree(os.path.join(os.path.split(docxfile)[0], 'temp_meta_docx/'))
