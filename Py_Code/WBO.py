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

# Create web browse off-line
class wboffline:
    
    def __init__(self,root):
        self.root = root
        root.title("Web Browsing Off-line")
        root.geometry("300x30+400+220")
        root.resizable(False,  False)
        self.root.bind('<Return>', self.gue)
        self.root.bind('<Control-P>', self.topdf)
        self.root.bind('<Control-p>', self.topdf)

        self.menu_bar = Menu(root)  # menu begins
        self.file_menu = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label='File', menu=self.file_menu)
        self.root.config(menu=self.menu_bar)  # menu ends
        
        self.file_menu.add_command(label='Save to PDF',  compound='left', 
                                   accelerator='Ctrl+P', command=self.topdf)

        self.st1 = StringVar()       
        self.fra1 = Frame(root)
        self.fra1.pack(pady=3)
        self.lab1 = Label(self.fra1, text="Url:")
        self.lab1.pack(side = LEFT)
        self.ent1 = Entry(self.fra1, textvariable=self.st1, width = 37)
        self.ent1.pack(side = LEFT)
        self.ent1.focus()
        self.but1 = Button(self.fra1, text="Save", command = self.gue)
        self.but1.pack(side = LEFT, padx=2)
        
        self.h_path = str(Path.home()) + "\\Documents"
        
        # Change Icon of tkinter
        self.root.iconbitmap(self.h_path + "\\WBO.ico")

    # Get the page and saved to a html file in computer
    def gue(self, event = None):
        try:
            ck = self.st1.get()
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
            on = str(Path.home())
            os.chdir(on)
            with open('wbo1.html','w') as sw:
                sw.write(str(text))
                sw.close
            on = str(Path.home())+ '\\wbo1.html'  
            webbrowser.open('file://' + on)
        except:
            mes.showerror('Error', sys.exc_info()[0])
    
    # Save to PDF and view it in browser
    def topdf(self, event = None):
        try:
            config = pdfkit.configuration(wkhtmltopdf='C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')
            options = {'quiet': ''}
            with open(str(Path.home())+ '\\out.pdf','w') as trun:
                trun.truncate()
                trun.close()
                
            ck = self.st1.get()
            if str(ck)[:8] != 'https://':
                pdfkit.from_url('https://' + ck, str(Path.home()) + '\\out.pdf',
                                configuration=config, options = options)
            else:
                pdfkit.from_url(ck, str(Path.home())+ '\\out.pdf', 
                                configuration=config, options = options)

        except:
            mes.showerror('Error', sys.exc_info()[0])
        finally:
            on = str(Path.home())+ '\\out.pdf'  
            webbrowser.open('file://' + on)    
            
begin = Tk()
my_gui = wboffline(begin)
begin.mainloop()


