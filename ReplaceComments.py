# -*- coding: utf-8 -*-
"""

@author: Maciej Ulaszewski
mail:ulaszewski.maciej@gmail.com
github: https://github.com/ulaszewskim

"""

import os
from CreateNewDocx import CreateNewDocx
from CreateTempZipDir import CreateTempZipDir
from shutil import rmtree

#==========
#Replace comments in entry docx file and save to other docx file
#==========
def ReplaceComments(coms_to_change, newfile, docxfile):
    CreateTempZipDir(os.path.split(docxfile)[1], os.path.split(docxfile)[0])
    
    xmlcom = open(os.path.join(os.path.split(docxfile)[0],'temp_meta_docx/word/comments.xml'), 'r').read()
    for i in range(len(coms_to_change)):
        xmlcom=xmlcom.replace(coms_to_change[i][0], coms_to_change[i][1])
        
    text_file = open(os.path.join(os.path.split(docxfile)[0],'temp_meta_docx/word/comments.xml'), "w")
    text_file.write(xmlcom)
    text_file.close()
    
    if os.path.exists(newfile+'.zip'):
        os.remove(newfile+'.zip')
    
    CreateNewDocx(os.path.join(os.path.split(docxfile)[0],'temp_meta_docx/'), newfile)
    rmtree(os.path.join(os.path.split(docxfile)[0],'temp_meta_docx/'))
    