# -*- coding: utf-8 -*-
"""
 
@author: Maciej Ulaszewski
mail:ulaszewski.maciej@gmail.com
github: https://github.com/ulaszewskim

"""

#==========
#get id, author, date and initials of comments 
#returns list of tuples, for example: ('1', 'Author Name', '2019-01-01T00:00:00Z', 'AN')
#==========
def get_comment_info(comments):
    ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    com=[]
    for comment in comments:
        idcom = str(comment.xpath('@w:id', namespaces=ns))[2:-2]
        author=str(comment.xpath('@w:author',namespaces=ns))[2:-2]
        date = str(comment.xpath('@w:date', namespaces=ns))[2:-2]
        initials = str(comment.xpath('@w:initials',namespaces=ns))[2:-2]

        com.append((idcom, author, date, initials,comment))
    return com
