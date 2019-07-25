# -*- coding: utf-8 -*-
"""

@author: Maciej Ulaszewski
mail:ulaszewski.maciej@gmail.com
github: https://github.com/ulaszewskim

"""

#==========
#create list of comments to change
#==========
def comments_to_change(cominfo, authors_merged):
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
