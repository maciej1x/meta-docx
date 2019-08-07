# -*- coding: utf-8 -*-
"""

@author: Maciej Ulaszewski
mail:ulaszewski.maciej@gmail.com
github: https://github.com/ulaszewskim

"""

import os
import zipfile


def create_new_docx(directory, docxname):
    """Create docx from directory"""
    zip_file = zipfile.ZipFile(docxname, 'w', zipfile.ZIP_DEFLATED)
    rootdir = os.path.basename(directory)
    for dirpath, dirnames, filenames in os.walk(directory):
        for filename in filenames:
            filepath = os.path.join(dirpath, filename)
            fullpath = os.path.relpath(filepath, directory)
            arcname = os.path.join(rootdir, fullpath)
            zip_file.write(filepath, arcname)
    zip_file.close()
