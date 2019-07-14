# -*- coding: utf-8 -*-
"""

@author: Maciej Ulaszewski
mail:ulaszewski.maciej@gmail.com
github: https://github.com/ulaszewskim

"""

import os
import zipfile
from shutil import copy2, rmtree

#==========
#create docx with new comments
#==========

def create_temp_zip_dir(docxfile, path_to_docxfile):
    if os.path.exists(os.path.join(path_to_docxfile,'temp_meta_docx')):
        rmtree(os.path.join(path_to_docxfile,'temp_meta_docx'))
    os.mkdir(os.path.join(path_to_docxfile,'temp_meta_docx'))
    
    zip_file_path = os.path.join(path_to_docxfile,'temp_meta_docx',docxfile[:-4]+'zip')
    copy2(os.path.join(path_to_docxfile,docxfile), zip_file_path)
    zip_file = zipfile.ZipFile(zip_file_path, 'r')
    zip_file.extractall(os.path.join(path_to_docxfile,'temp_meta_docx/'))
    zip_file.close()
    os.remove(zip_file_path)
