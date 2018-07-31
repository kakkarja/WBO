# -*- coding: utf-8 -*-
"""
Created on Fri Apr 13 11:56:59 2018

@author: karja
"""
from tkinter import *
import tkinter.messagebox as mes
import requests
from bs4 import BeautifulSoup
import webbrowser,os
from pathlib import Path
import sys
import pdfkit
import time

# Create web browse off-line
class wboffline:
    
    def __init__(self,root):
        self.root = root
        root.title("Web Browsing Off-line")
        root.geometry("300x30+450+220")
        root.resizable(False,  False)
        self.root.bind('<Return>', self.wbp)
        self.root.bind('<Control-P>', self.topdf)
        self.root.bind('<Control-p>', self.topdf)

        self.menu_bar = Menu(root)  # menu begins
        self.file_menu = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label='File', menu=self.file_menu)
        self.root.config(menu=self.menu_bar)  # menu ends
        
        self.file_menu.add_command(label='Save to PDF',  compound='left', 
                                   accelerator='Ctrl+P', command=self.topdf)
        
        self.about_menu = Menu(self.menu_bar, tearoff = 0)
        self.menu_bar.add_cascade(label = 'About', menu = self.about_menu)
        self.about_menu.add_command(label = 'Help',compound='left', 
                                    command=self.about)

        self.st1 = StringVar()       
        self.fra1 = Frame(root)
        self.fra1.pack(pady=3)
        self.lab1 = Label(self.fra1, text="Url:")
        self.lab1.pack(side = LEFT)
        self.ent1 = Entry(self.fra1, textvariable=self.st1, width = 37)
        self.ent1.pack(side = LEFT)
        self.ent1.focus()
        self.but1 = Button(self.fra1, text="Save", command = self.wbp)
        self.but1.pack(side = LEFT, padx=3)
        
        self.h_path = str(Path.home()) + "\\Documents"

        try:
            self.h_path = str(Path.home()) + "\\Documents\\WBO_Files"
            ch = os.chdir(self.h_path)
        except:
            os.mkdir(self.h_path)
            ms = 'has been created.'
            mes.showinfo('WBO', 'The path to WBO_Files, {}'.format(ms))
        
        # Change Icon of tkinter
        dir_p ='C:\\Program Files\\K A K\\WBO'
        self.root.iconbitmap(dir_p + "\\WBO.ico")

    # Get the page and saved to a html file in computer
    def wbp(self, event = None):
        ck = self.st1.get()
        try:
            if str(ck)[:8] != 'https://':
                with requests.Session() as s:
                    req = s.get('https://' + ck)
            else:
                with requests.Session() as s:
                    req = s.get(ck)
                    
            soup = BeautifulSoup(req.text, 'html.parser')
            text = str(soup).replace('\n',"").replace('\t',
                    "").replace('\r',"").replace('href=',
                    "").replace('action=',"").replace('image='," ").encode('ascii',
                    'xmlcharrefreplace')
            s_name = time.ctime(time.time()).replace(' ','').replace(':','') + '_' + ck + '.html'
            with open(s_name,'w') as sw:
                sw.write(str(text))
                sw.close
            on = self.h_path + '\\' + s_name
            os.startfile(on)
        except:
            mes.showerror('Error', sys.exc_info()[0:])
    
    '''
    # Save to PDF and view it in browser. Please download
    # wkhtmltopdf.exe file in:
    # https://wkhtmltopdf.org/downloads.html
    # to be able to use this function.
    '''
    def topdf(self, event = None):
        config = pdfkit.configuration(wkhtmltopdf='C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')
        options = {'quiet': ''}
        ck = self.st1.get()
        s_name = time.ctime(time.time()).replace(' ','').replace(':','') + '_' + ck + '.pdf'
        with open(s_name,'w') as trun:
            trun.truncate()
            trun.close()
        try:
            if str(ck)[:8] != 'https://':
                pdfkit.from_url('https://' + ck, s_name,
                                configuration=config, options = options)
            else:
                pdfkit.from_url(ck, s_name, 
                                configuration=config, options = options)
        except:
            mes.showerror('Error', sys.exc_info()[0:])
        finally:
            on = self.h_path + '\\' + s_name
            os.startfile(on)  
            
    # Link to WBO Github page
    def about(self):
        webbrowser.open_new(r"https://github.com/kakkarja/WBO")
        
if __name__ == '__main__':            
    begin = Tk()
    my_gui = wboffline(begin)
    begin.mainloop()