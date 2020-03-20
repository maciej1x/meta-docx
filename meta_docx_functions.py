# -*- coding: utf-8 -*-
"""

@author: Maciej Ulaszewski
mail:ulaszewski.maciej@gmail.com
github: https://github.com/ulaszewskim

"""
import os
import zipfile
from shutil import copy2, rmtree
from lxml import etree


def get_unique_authors(comments):
    """
    get all unique authors names with initials of comments
    returns list of tuples
    for example: ('Author Name','AN')
    """
    ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    authors = []
    for comment in comments:
        author = str(comment.xpath('@w:author', namespaces=ns))[2:-2]
        initials = str(comment.xpath('@w:initials', namespaces=ns))[2:-2]
        authors.append((author, initials))
    authors = list(dict.fromkeys(authors)) #remove duplicates
    return authors

def get_comment_elements(docxfile):
    """get comments elements from docx file"""
    ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    docx_zip = zipfile.ZipFile(docxfile)
    comments_xml = docx_zip.read('word/comments.xml')
    et = etree.XML(comments_xml)
    comments = et.xpath('//w:comment', namespaces=ns)
    return comments

def get_comment_info(comments):
    """
    get id, author, date and initials of comments 
    returns list of tuples
    for example: ('1', 'Author Name', '2019-01-01T00:00:00Z', 'AN')
    """
    ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    com = []
    for comment in comments:
        idcom = str(comment.xpath('@w:id', namespaces=ns))[2:-2]
        author = str(comment.xpath('@w:author', namespaces=ns))[2:-2]
        date = str(comment.xpath('@w:date', namespaces=ns))[2:-2]
        initials = str(comment.xpath('@w:initials', namespaces=ns))[2:-2]

        com.append((idcom, author, date, initials, comment))
    return com

def create_temp_zip_dir(docxfile, path_to_docxfile):
    """create docx with new comments"""
    if os.path.exists(os.path.join(path_to_docxfile, 'temp_meta_docx')):
        rmtree(os.path.join(path_to_docxfile, 'temp_meta_docx'))
    os.mkdir(os.path.join(path_to_docxfile, 'temp_meta_docx'))
    zip_file_path = os.path.join(path_to_docxfile, 'temp_meta_docx', docxfile[:-4]+'zip')
    copy2(os.path.join(path_to_docxfile, docxfile), zip_file_path)
    zip_file = zipfile.ZipFile(zip_file_path, 'r')
    zip_file.extractall(os.path.join(path_to_docxfile, 'temp_meta_docx/'))
    zip_file.close()
    os.remove(zip_file_path)

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

def comments_to_change(cominfo, authors_merged):
    """create list of comments to change"""
    coms_to_change = []
    for c in range(len(cominfo)):
        for a in range(len(authors_merged)):
            if cominfo[c][1] == authors_merged[a][0] and cominfo[c][3] == authors_merged[a][1]:
                if cominfo[c][2] == '':
                    original = '<w:comment w:id="'+cominfo[c][0]+'" w:author="'+cominfo[c][1]+'" w:initials="'+cominfo[c][3]+'">'
                    changed = '<w:comment w:id="'+cominfo[c][0]+'" w:author="'+authors_merged[a][2]+'" w:initials="'+authors_merged[a][3]+'">'
                    coms_to_change.append([original, changed])
                else:
                    original = '<w:comment w:id="'+cominfo[c][0]+'" w:author="'+cominfo[c][1]+'" w:date="'+cominfo[c][2]+'" w:initials="'+cominfo[c][3]+'">'
                    changed = '<w:comment w:id="'+cominfo[c][0]+'" w:author="'+authors_merged[a][2]+'" w:date="'+cominfo[c][2]+'" w:initials="'+authors_merged[a][3]+'">'
                    coms_to_change.append([original, changed])
    return coms_to_change

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
