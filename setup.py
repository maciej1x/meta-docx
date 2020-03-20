# -*- coding: utf-8 -*-
"""

@author: Maciej Ulaszewski
mail:ulaszewski.maciej@gmail.com
github: https://github.com/ulaszewskim

"""


import os
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as msg
from meta_docx_functions import *


#===================
#GUI
#===================

#==========
#GUI functions
#==========


def select_source_docx():
    """Select source docx file"""
    filename = fd.askopenfilename(filetypes=[('.docx', '*.docx')])
    if filename != '':
        src_dir.delete(0, tk.END)
        src_dir.insert(tk.INSERT, filename)
        dst_dir['state'] = 'normal'
        dst_dir.update()
        dst_btn['state'] = 'normal'
        dst_btn.update()
        dst_dir.delete(0, tk.END)
        dst_dir.insert(tk.INSERT, filename[:-5]+"_new.docx")
        start_btn['state'] = 'normal'
        start_btn['bg'] = 'lightgreen'
        start_btn.update()
    fill_authors(filename)



def select_new_docx():
    """Select new docx file directory"""
    sourcefile = dst_dir.get()
    newfile = fd.askdirectory()
    if newfile != '':
        dst_dir.delete(0, tk.END)
        dst_dir.insert(tk.INSERT, newfile+'/'+os.path.basename(sourcefile))
    newfile = dst_dir.get()
    return newfile #returns full path to new file



def check():
    """Reload entries when any checkbox changes its value"""
    i = 2
    checks = []
    while True:
        try:
            btn = comments.grid_slaves(column=0, row=i)[0]
            a = btn.configure()['variable'][4]
            b = int(btn.getvar(name=a))
            checks.append(b)
            if b == 0:
                old_author = comments.grid_slaves(column=1, row=i)[0]
                old_author.config(disabledbackground='gray95')
                old_author.update()
                old_initial = comments.grid_slaves(column=2, row=i)[0]
                old_initial.config(disabledbackground='gray95')
                old_initial.update()
                new_author = comments.grid_slaves(column=3, row=i)[0]
                new_author.config(state='disabled', disabledforeground='gray95')
                new_author.update()
                new_initial = comments.grid_slaves(column=4, row=i)[0]
                new_initial.config(state='disabled', disabledforeground='gray95')
                new_initial.update()
            else:
                old_author = comments.grid_slaves(column=1, row=i)[0]
                old_author.config(disabledbackground='DarkSeaGreen1')
                old_author.update()
                old_initial = comments.grid_slaves(column=2, row=i)[0]
                old_initial.config(disabledbackground='DarkSeaGreen1')
                old_initial.update()
                new_author = comments.grid_slaves(column=3, row=i)[0]
                new_author.config(state='normal')
                new_author.update()
                new_initial = comments.grid_slaves(column=4, row=i)[0]
                new_initial.config(state='normal')
                new_initial.update()
            i = i+1
        except IndexError:
            break
    return checks



def fill_authors(filename):
    """Fill entries with authors"""
    comments1 = get_comment_elements(filename)
    authors = get_unique_authors(comments1)
    row = 2
    for author in authors:
        src_author = author[0]
        src_initial = author[1]
        var = tk.StringVar()
        var.set(src_author)
        old_author = tk.Entry(comments, textvariable=var, state='disabled', disabledbackground="gray95", disabledforeground='black', width=35)
        old_author.grid(row=row, column=1, sticky='W', padx=2)
        var = tk.StringVar()
        var.set(src_initial)
        old_initial = tk.Entry(comments, textvariable=var, state='disabled', disabledbackground="gray95", disabledforeground='black', width=7)
        old_initial.grid(row=row, column=2, sticky='W', padx=2)
        new_author_var = tk.StringVar()
        new_author = tk.Entry(comments, textvariable=new_author_var, width=35, state='disabled')
        new_author.grid(row=row, column=3, sticky='W', padx=2)
        new_initial_var = tk.StringVar()
        new_initial = tk.Entry(comments, textvariable=new_initial_var, width=7, state='disabled')
        new_initial.grid(row=row, column=4, sticky='W', padx=2)
        chk = tk.Checkbutton(comments, command=check)
        chk.grid(row=row, column=0)
        row += 1



def start():
    """Main function"""
    chk_var = check() #get checkboxes values
    fulllen = len(chk_var) #number of authors
    #call functions needed to replace comments
    docxfile = src_dir.get()
    comments_docx = get_comment_elements(docxfile)
    authors = get_unique_authors(comments_docx)
    cominfo = get_comment_info(comments_docx)
    #create tuple with new authors
    authors_to_change = []
    for i in range(fulllen):
        author = comments.grid_slaves(column=3, row=i+2)[0].get()
        initial = comments.grid_slaves(column=4, row=i+2)[0].get()
        authors_to_change.append((author, initial))
    #create tuple with old authors
    authors_merged = []
    for j in range(fulllen):
        authors_merged.append(authors[j]+authors_to_change[j])
    #remove unwanted changes
    wanted_changes = []
    for k in range(fulllen):
        if chk_var[k] == 1:
            wanted_changes.append(authors_merged[k])
    coms_to_change = comments_to_change(cominfo, wanted_changes)
    newfile = dst_dir.get()
    #replace comments
    replace_comments(coms_to_change, newfile, docxfile)
    msg.showinfo('Finished successfully', 'New file: {}'.format(newfile))

if __name__ == "__main__":
    #==========
    #Creating a window
    #==========
    window = tk.Tk()
    window.title('Reviewer editor')
    window.geometry('600x600+20+20')

    frame1 = tk.Frame(window)
    frame1.grid(sticky='news')


    #Start button
    start_btn = tk.Button(frame1, text="Start", width=64, font=('', '12'), state='disabled', command=start)
    start_btn.grid(column=1, row=0, columnspan=5, sticky='W', padx=5, pady=5)

    #Source file
    src_h = tk.Label(frame1, text="Select file to edit:")
    src_h.grid(column=1, row=1, sticky='W', padx=5)

    src_docx = tk.StringVar() #source docx file
    src_dir = tk.Entry(frame1, width=88, state='normal', textvariable=src_docx)
    src_dir.grid(column=1, row=2, columnspan=4, sticky='W', padx=5, pady=0)

    src_btn = tk.Button(frame1, text='Select', command=select_source_docx)
    src_btn.grid(column=5, row=2, sticky='E')

    #Destination file
    dst_h = tk.Label(frame1, text="New file directory:")
    dst_h.grid(column=1, row=3, sticky='W', padx=5)

    dst_docx = tk.StringVar() #new docx directory
    dst_dir = tk.Entry(frame1, width=88, state='disabled', textvariable=dst_docx)
    dst_dir.grid(column=1, row=4, columnspan=4, sticky='W', padx=5, pady=0)

    dst_btn = tk.Button(frame1, state='disabled', text='Select', command=select_new_docx)
    dst_btn.grid(column=5, row=4, sticky='E')


    #=======
    #Comments 
    #=======
    frame_canvas = tk.Frame(frame1)
    frame_canvas.grid(row=6, column=1, columnspan=6, sticky='news', pady=5, padx=5)
    frame_canvas.grid_propagate(True)

    canvas = tk.Canvas(frame_canvas, height=450, width=570)
    canvas.grid(row=0, column=0, sticky="news")

    #Scrollbar
    vsb = tk.Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
    vsb.grid(row=0, column=1, sticky='ns')
    canvas.configure(yscrollcommand=vsb.set)

    comments = tk.Frame(canvas)
    canvas.create_window((0, 0), window=comments, anchor='nw')

    #Headers
    old_h = tk.Label(comments, text="Authors in source file:")
    old_h.grid(column=1, row=0, sticky='W')
    old_hn = tk.Label(comments, text="Name")
    old_hn.grid(column=1, row=1, sticky='W')
    old_hi = tk.Label(comments, text="Initials")
    old_hi.grid(column=2, row=1, sticky='W')
    new_h = tk.Label(comments, text="New authors:")
    new_h.grid(column=3, row=0, sticky='W')
    new_hn = tk.Label(comments, text="Name")
    new_hn.grid(column=3, row=1, sticky='W')
    new_hi = tk.Label(comments, text="Initials")
    new_hi.grid(column=4, row=1, sticky='W')

    #scroll options
    comments.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    window.mainloop()
