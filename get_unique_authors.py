# -*- coding: utf-8 -*-
"""

@author: Maciej Ulaszewski
mail: ulaszewski.maciej@gmail.com
github: https://github.com/ulaszewskim

"""

#==========
#get all unique authors names with initials of comments
#returns list of tuples, for example: ('Author Name','AN')
#==========
def get_unique_authors(comments):
    ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    authors=[]
    for comment in comments:
        author=str(comment.xpath('@w:author',namespaces=ns))[2:-2]
        initials = str(comment.xpath('@w:initials',namespaces=ns))[2:-2]
        authors.append((author, initials))
    authors = list(dict.fromkeys(authors)) #remove duplicates
    return authors
