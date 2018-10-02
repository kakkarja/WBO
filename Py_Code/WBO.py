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
import re
import win32com.client as win32
import sys
import tkinter.simpledialog as sim

# Create web browse off-line
class wboffline:
    def __init__(self,root):
        self.root = root
        root.title("Web Browsing Off-line")
        root.geometry("300x30+450+220")
        root.resizable(False,  False)
        root.attributes('-topmost', True)
        self.root.bind('<Return>', self.wbp)
        self.root.bind('<Control-P>', self.topdf)
        self.root.bind('<Control-p>', self.topdf)
        self.root.bind('<Control-T>', self.totxt)
        self.root.bind('<Control-t>', self.totxt)

        self.menu_bar = Menu(root)  # menu begins
        self.file_menu = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label='File', menu=self.file_menu)
        self.root.config(menu=self.menu_bar)  # menu ends
        
        self.file_menu.add_command(label='Save to PDF',  compound='left', 
                                   accelerator='Ctrl+P', command=self.topdf)
        self.file_menu.add_command(label='Text to PDF',  compound='left', 
                                   accelerator='Ctrl+T', command=self.totxt)
        
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
        self.ent1.bind('<Enter>', self.pas)
        self.ent1.pack(side = LEFT)
        self.ent1.focus()
        self.but1 = Button(self.fra1, text="Save", command = self.wbp)
        self.but1.pack(side = LEFT, padx=3)
        
        self.h_path = str(Path.home()) + "\\Documents"

        try:
            self.h_path = str(Path.home()) + "\\Documents\\WBO_Files"
            os.chdir(self.h_path)
        except:
            os.mkdir(self.h_path)
            ms = 'has been created.'
            mes.showinfo('WBO', 'The path to WBO_Files, {}'.format(ms))
            os.chdir(self.h_path)
        
        # Change Icon of tkinter
        dir_p = os.path.dirname(os.path.realpath('WBO.ico'))
        try:
            self.root.iconbitmap(dir_p + "\\WBO.ico")
        except:
            pass

    # Changing url string to text name
    def exu(self,url): 
        regex = r'((\w+\.)+(\W\w+|\w+)+)'
        fix = re.sub(r'\W', '_', str(re.search(regex,url).group()))
        if 'www' in fix:
            fix = fix[4:]
        return fix   

    # Get the page and saved to a html file in computer
    def wbp(self, event = None):
        ck = self.st1.get()
        if '\r' in ck:
            ck = ck.replace('\r','')
        try:
            st_u = self.exu(ck)
            if 'https://' not in ck:
                with requests.Session() as s:
                    req = s.get('https://' + ck)
            else:
                with requests.Session() as s:
                    req = s.get(ck)
            soup = BeautifulSoup(req.text, 'html.parser')
            text = re.sub('href=|action=', '', str(soup)).encode('ascii', 'xmlcharrefreplace').decode('raw_unicode_escape')
            s_name = time.ctime(time.time()).replace(' ','').replace(':','') + '_' + st_u + '.html'
            with open(s_name,'w') as sw:
                sw.write(str(text))
                sw.close
            on = self.h_path + '\\' + s_name
            os.startfile(on)
        except:
            mes.showerror('Error', sys.exc_info()[1])
    
    '''
    # Save to PDF and view it in browser. Please download
    # wkhtmltopdf.exe file in:
    # https://wkhtmltopdf.org/downloads.html
    # to be able to use this function.
    '''
    def topdf(self, event = None):
        pat = 'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'
        config = pdfkit.configuration(wkhtmltopdf = pat)
        options = {'quiet': ''}
        ck = self.st1.get()
        if '\r' in ck:
            ck = ck.replace('\r','')
        try:
            st_u = self.exu(ck)
            s_name = time.ctime(time.time()).replace(' ','').replace(':','') + '_' + st_u + '.pdf'
            with open(s_name,'w') as trun:
                trun.truncate()
                trun.close()
            if 'https://' not in ck:
                pdfkit.from_url('https://' + ck, s_name,
                                configuration=config, options = options)
            else:
                pdfkit.from_url(ck, s_name, 
                                configuration=config, options = options)
        except:
            mes.showerror('Error', sys.exc_info()[1])
        finally:
            try:
                on = self.h_path + '\\' + s_name
                os.startfile(on)
            except:
                pass

    # Text to Pdf
    def totxt(self, event = None):
        ck = self.ent1.get()
        if '\r' in ck:
            ck = ck.replace('\r','')
        else:
            chan = ck[-1]+ck[-2]
            ck = ck[:-2]+chan
            self.ent1.delete(0, 'end')
            self.ent1.insert(0, ck)
        try:
            #st_u = self.exu(ck)
            if 'http://' in ck:
                with requests.Session() as s:
                    req = s.get(ck)
            elif 'https://' not in ck:
                with requests.Session() as s:
                    req = s.get('https://' + ck)
            else:
                with requests.Session() as s:
                    req = s.get(ck)
            soup = BeautifulSoup(req.text, 'html.parser')
            for script in soup(["script", "style"]):
                script.extract()  
            text = soup.get_text()
            lines = (line.strip() for line in text.splitlines())
            chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
            txt_web = ""
            while True:
                try: 
                    txt_web += ''.join(next(chunks))+'\n'
                except:
                    break
            txt_web = re.sub(r'(\n\n\n)','',txt_web)
        except:
            mes.showerror('Not Secure', sys.exc_info()[1])
        else:
            file_n =  sim.askstring('Save as PDF', 'Input file name:')
            try:
                word = win32.gencache.EnsureDispatch('Word.Application')
                doc = word.Documents.Add()
                rng = doc.Range(0,0)
                rng.InsertAfter(txt_web)
                doc.SaveAs(file_n, FileFormat = 17)
                doc.Close(False)
                word.Application.Quit()
                os.rename('{}\\Documents\\{}.pdf'.format(str(Path.home()), file_n), '{}\\Documents\\WBO_Files\\{}.pdf'.format(str(Path.home()), file_n))
                os.startfile('{}\\Documents\\WBO_Files\\{}.pdf'.format(str(Path.home()), file_n))
            except:
                mes.showerror('Not Secure', sys.exc_info()[1])
    
    # Self paste from the clipboard
    def pas(self, event):
        try:
            text = self.root.clipboard_get()
            ck = len(self.ent1.get())
        except:
            pass
        else:
            if ck <= 1:
                self.ent1.insert(0, text)
                self.root.clipboard_clear()
            else:
                aq = mes.askquestion('Clear', 'Do you want to clear the text?')
                if aq == 'yes':
                    self.ent1.delete(0, 'end')
                    self.ent1.insert(0, text)
                    self.root.clipboard_clear()
                else:
                    self.root.clipboard_clear()

    # Link to WBO Github page
    def about(self):
        webbrowser.open_new(r"https://github.com/kakkarja/WBO")

if __name__ == '__main__':         
    begin = Tk()
    my_gui = wboffline(begin)
    begin.mainloop()