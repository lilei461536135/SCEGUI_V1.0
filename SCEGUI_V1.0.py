# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import *
from tkinter import ttk
import threading
import time
import re
import os
import sys
import queue
import subprocess
import multiprocessing
from tkinter.messagebox import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

try:
    file_dir = os.path.dirname(os.path.abspath(__file__))
except NameError:  # We are the main py2exe script, not a module
    file_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    
os.chdir(file_dir)

# common definitions
sce_tool = "%s\\SCEWIN_64.exe" % file_dir
# exported bios file
bios_exported_file = "%s\\bios_e.txt" % file_dir
# export bios file log
export_log = "%s\\export.log" % file_dir
# file to be import
bios_import_file = "%s\\bios_i.txt" % file_dir
# import bios file log
import_log = "%s\\import.log" % file_dir

# self defined entry class
class EntryClass(ttk.Entry):
    def __init__(self, master, width, row, column, state="normal", string=""):
        ttk.Entry.__init__(self, master=master, width=width, state=state)
        self.config(font =("Consolas", 12, "normal"))
        # content in the entry object
        self.string = tk.StringVar()
        self.string.set(string)
        self.configure(textvariable=self.string)
        self.grid(row=row, column=column, sticky=tk.W)

    # get string in entry box
    def get_string(self):
        text = self.string.get()
        return text

    # set content
    def set_string(self, string):
        self.string.set(string)

    # clear content
    def clear(self):
        self.string.set("")


# self defined ComboBox Class
class ComboBoxClass(ttk.Combobox):
    def __init__(self, master, width, row, column, column_span=1, sticky=tk.W):
        ttk.Combobox.__init__(self, master=master, width=width)
        self.config(font =("Consolas", 12, "normal"))
        self.string = tk.StringVar()
        self.configure(textvariable=self.string)
        self.grid(row=row, column=column, columnspan=column_span, sticky=sticky)

    # get content showing in the box
    def get_string(self):
        return self.get()

    # set content showing in the box
    def set_string(self, string):
        self.set(string)

    # set selection list
    def set_list(self, data_dict):
        self["value"] = data_dict

    # clear box
    def clear(self):
        self["value"] = ""
        self.delete(0, len(self.get_string()))


# self defined Button Class
class ButtonClass(tk.Button):
    def __init__(self, master, label, row, column, column_span=1, width=None, fg=None, bg="gray", state="normal",
                 height=1, row_span=1, sticky=tk.W, activebg=None, command=None):
        tk.Button.__init__( \
            self, master=master, text=label, width=width, \
            height=height, fg=fg, bg=bg, state=state, activebackground=activebg, command=command)
        self.config(font =("Consolas", 12, "normal"))
        self.grid(row=row, column=column, columnspan=column_span, rowspan=row_span, sticky=sticky)


# self defined Label Class
class LabelClass(ttk.Label):
    def __init__(self, master, string, row, column, column_span=1, width=10, fg="black"):
        ttk.Label.__init__(self, master=master, text=string, width=width, foreground=fg)
        self.config(font =("Consolas", 12, "normal"))
        self.grid(row=row, column=column, columnspan=column_span, sticky=tk.W)

# self defined Text Class
class TextClass(tk.Text):
    def __init__(self, master, row, column, width, height, column_span=1):
        tk.Text.__init__(self, master=master, width=width, height=height)
        self.config(font =("Consolas", 12, "normal"))
        self.scrollbar = tk.Scrollbar(master=master)
        self.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.yview)
        self.grid(row=row, column=column, columnspan=column_span, sticky=tk.W)
        self.scrollbar.grid(row=row, column=column, columnspan=column_span, sticky=tk.E+tk.N+tk.S)
        
    def clear(self):
        self.delete('0.0', END)

# self defined ScrolledList Class
class ScrolledListClass:
    def __init__(self, master, width, row, column, height=1, \
                 row_span=1, column_span=1):
        # content of selected item
        self.selection = None
        # all contents in listbox
        self.contents = None
        # length of contents in listbox
        self.len = 0
        self.string = tk.StringVar()
        self.listbox = tk.Listbox(master=master, width=width, height=height,
                                  listvariable=self.string, selectmode="single", selectbackground="yellow", selectforeground="black")
        self.listbox.config(font =("Consolas", 12, "normal"))
        self.scrollbar = tk.Scrollbar(master=master)
        self.listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.listbox.yview)
        self.listbox.grid(row=row, column=column, rowspan=row_span, \
                          columnspan=column_span, sticky=tk.W)
        self.scrollbar.grid(row=row, column=column, rowspan=row_span, \
                            columnspan=column_span, sticky=tk.E+tk.N+tk.S)

    # function to get current selected item
    def get_selection(self):
        self.listbox.update()
        self.selection = self.listbox.get(ANCHOR)

    # insert string to listbox
    def insert(self, pos, string):
        self.listbox.insert(pos, string)
        self.len += 1

    # delete an item at specified position
    def delete(self, pos):
        self.listbox.delete(pos)
        self.len -= 1

    # clear all items in listbox
    def clear(self):
        size = self.listbox.size()
        self.listbox.delete(0, size)
        self.len = 0

    # get all contents in listbox
    def get(self):
        size = self.listbox.size()
        self.contents = self.listbox.get(0, size)
        return self.contents

    # show index line
    def see(self, index):
        self.listbox.see(index)
    

# GUI main window class
class GUIMainWin:
    def __init__(self):
        self.bios_dict = {}  # dict with all bios settings
        self.bios_content = None
        # a queue to store matched items in bios
        self.matched_dataQueue = queue.Queue()
        self.root = tk.Tk()
        self.root.title("SCEGUI")
        self.root.resizable(False, False)
        # Add accelerator
        self.root.bind_all('<Control-o>', self.event_open)
        self.root.bind_all('<Control-s>', self.event_save)
        self.root.bind_all('<Control-i>', self.event_import)
        self.root.bind_all('<Control-r>', self.event_default)
        # Add menu
        menubar = tk.Menu(self.root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Open", command=self.callback_open, accelerator='Ctl+o')
        filemenu.add_separator()
        filemenu.add_command(label="Save as", command=self.callback_save, accelerator='Ctl+s')
        menubar.add_cascade(label="File", menu=filemenu)
        actionmenu = tk.Menu(menubar, tearoff=0)
        actionmenu.add_command(label="Import to BIOS NVRAM", command=self.callback_import, accelerator='Ctl+i')
        actionmenu.add_separator()
        actionmenu.add_command(label="Restore Defaults", command=self.callback_default, accelerator='Ctl+r')
        menubar.add_cascade(label="Action", menu=actionmenu)
        self.root.config(menu=menubar)
        # Pop menu
        self.postmenu = tk.Menu(self.root)
        self.postmenu_x = 0
        self.postmenu_y = 0
        # Add components
        self.text_bios = TextClass(master=self.root, row=0, column=0, width=80, height=20, column_span=4)
        self.label_search = LabelClass( \
            master=self.root, string="Search :", \
            row=1, column=0, width=10, fg="blue")
        self.entry_search = EntryClass( \
            master=self.root, width=25, row=1, column=1)
        self.entry_search.bind('<KeyRelease>', self.event_search_KeyRelease)
        self.label_select = LabelClass( \
            master=self.root, string="Select(or)Input :", \
            row=1, column=2, width=15, fg="blue")
        self.combobox_value = ComboBoxClass( \
            master=self.root, width=20, row=1, column=3, sticky=tk.W)
        self.button_update = ButtonClass( \
            master=self.root, label="Update", row=2, row_span=2, column=3, \
            width=20, height=3, sticky=tk.W, activebg='green', command=self.callback_update)
        self.label_match = LabelClass( \
            master=self.root, string="Options :", \
            row=2, column=0, width=10, fg="black")
        self.list_match = ScrolledListClass( \
            master=self.root, width=25, height=4, row=2, column=1)
        # self.list_match.listbox.bind('<Button-1>', self.event_match_B1)
        self.list_match.listbox.bind('<ButtonRelease-1>', self.event_match_B1Release)
        self.label_input = LabelClass( \
            master=self.root, string="Message :", \
            row=3, column=0, width=10, fg="black")
        self.info_list = ScrolledListClass( \
            master=self.root, width=80, height=5, \
            row=4, column=0, column_span=4)
        self.root.update()
        threading.Thread(target=self.initialize).start()  # Export BIOS from NVRAM at start
        self.root.mainloop()
        
    # export current bios settings
    def initialize(self):
        # Export BIOS settings from NVRAM at startup
        self.export_bios()
        # get bios settings
        with open(bios_exported_file, 'r') as f:
            self.bios_content = f.readlines()
            f.close()
        self.show_bios()
        
    # export current bios settings from NVRAM
    def export_bios(self):
        try:
            self.info_list.insert(0, "Exporting Current BIOS Settings, please wait...")
            command = "%s /o /s %s" % (sce_tool, bios_exported_file)
            with open(export_log, 'w') as fo:
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                process = subprocess.Popen(command,  stdout=fo, stderr=fo, stdin=subprocess.PIPE, startupinfo=si)
                process.wait()
                fo.close()
            # show sce export messages to listbox
            self.info_list.clear()
            with open(export_log, 'r') as fo:
                i = 0
                for line in fo.readlines():
                    self.info_list.insert(i, line)
                    i += 1
                fo.close()
        except Exception as e:
            showerror("Error", str(e))

    # function to make bios item - value dict
    def make_bios_dict(self):
        try:
            # clear old contents
            self.bios_dict.clear()
            pattern1 = re.compile(r"^Setup Question.*=", re.I)
            pattern2 = re.compile(r"^Options.*=", re.I)
            pattern3 = re.compile(r"^Value.*=", re.I)
            pattern4 = re.compile(r"^BIOS Default.*=", re.I)
            pattern5 = re.compile(r"\*\[.*\]", re.I)
            key_found = 0
            option_found = 0
            default_found = 0
            index = -1  # line num in self.bios_content
            for line in self.bios_content:
                index += 1
                # search "Options" keyword
                if key_found == 0 and line != "\n":
                    if re.search(pattern1, line):
                        key = line.split("=")[1].split("//")[0].strip()
                        if key in self.bios_dict.keys():
                            continue
                        key_found = 1
                        self.bios_dict[key] = {}
                        # store key index, if multi exists, store only first
                        self.bios_dict[key]['index'] = index
                        continue
                elif key_found == 1 and line == "\n":
                    key_found = 0
                    option_found = 0
                    default_found = 0
                    continue
                if key_found == 1:
                    # search "BIOS Default" keyword
                    if re.search(pattern4, line):
                        default_line = line
                        default_found = 1
                        continue
                    # search keyword Options
                    if re.search(pattern2, line):
                        if default_found == 1:
                            # default value store
                            default_value = default_line.split("=")[1].split("]")[1].split("//")[0].strip()
                            self.bios_dict[key]['default'] = default_value
                        else:
                            self.bios_dict[key]['default'] = 'empty'
                        option_found = 1
                        value = line.split("]")[1].split("//")[0].strip()
                        self.bios_dict[key].setdefault('options', []).append(value)
                        if re.search(pattern5, line):
                            self.bios_dict[key]['current'] = value
                        continue
                    if option_found == 1:
                        value = line.split("]")[1].split("//")[0].strip()
                        self.bios_dict[key].setdefault('options', []).append(value)
                        if re.search(pattern5, line):
                            self.bios_dict[key]['current'] = value
                        continue
                    if re.search(pattern3, line):
                        if default_found == 1:
                            # default value stored in array[0]
                            default_value = default_line.split("=")[1].split("//")[0].strip()
                            self.bios_dict[key]['default'] = default_value
                        else:
                            self.bios_dict[key]['default'] = 'empty'
                        value = line.split("=")[1].split("//")[0].strip()
                        self.bios_dict[key]['value'] = value
                        self.bios_dict[key]['current'] = value
                        continue
        except Exception as e:
            showerror("Error", str(e))

    # show bios settings in list_bios 
    def show_bios(self):
        self.text_bios.config(state='normal')
        self.text_bios.configure(bg='gray')
        self.text_bios.clear()
        for line in self.bios_content:
            self.text_bios.insert(INSERT, line)
        self.text_bios.configure(bg='white')
        self.make_bios_dict()
        self.text_bios.config(state='disabled')

    def update_one_bios(self, content, key, value):
        try:
            self.text_bios.config(state='normal')
            key = key.strip()
            value = value.strip()  # very important!!!
            pattern1 = re.compile(r"^Setup Question.*=")
            pattern2 = re.compile(r"^Options.*=")
            pattern3 = re.compile(r"^Value.*=")
            pattern4 = re.compile(r"\*\[.*\]")
            beg2modify = 0
            beg2modify_opts = 0
            self.list_match.listbox.configure(bg='gray')
            index = -1  # index for bios_content
            index2 = 0  # index for text_bios
            for line in content:
                index += 1
                index2 += 1
                if line != "\n":
                    # search identifier "Setup Question"
                    if re.search(pattern1, line) and beg2modify == 0:
                        # match wanted key with item
                        if line.split("=")[1].split("//")[0].strip() == key:
                            self.info_list.insert(self.info_list.len, "Begin to edit")
                            beg2modify = 1
                    if beg2modify == 1:
                        if re.search(pattern2, line):
                            beg2modify_opts = 1
                        elif re.search(pattern3, line):
                            new_line = "Value   =%s\n" % value
                            self.info_list.insert(self.info_list.len, "Change value")
                            content[index] = content[index].replace(line, new_line)
                            self.text_bios.see(str(index2)+'.0')
                            self.text_bios.delete(str(index2)+'.0', str(index2)+'.end')
                            self.text_bios.insert(str(index2)+'.0', new_line.strip('\n')) 
                            self.make_bios_dict()
                            self.text_bios.update()
                            self.root.update()  
                    if beg2modify_opts == 1:
                        # Options	=*[00]Performance	// Move "*" to the desired Option
                        #            [07]Balanced Performance
                        #            [08]Balanced Power
                        #            [0F]Power

                        # not found "*[" but found value
                        # get value
                        tmp_value = line.split("]")[1].split("//")[0].strip()
                        if re.search(pattern4, line) is None and (value == tmp_value):
                            self.info_list.insert(self.info_list.len, "Add *")
                            new_line = line.replace("[", "*[")
                            content[index] = content[index].replace(line, new_line)
                            self.text_bios.see(str(index2)+'.0')
                            self.text_bios.delete(str(index2)+'.0', str(index2)+'.end')
                            self.text_bios.insert(str(index2)+'.0', new_line.strip('\n'))
                            self.make_bios_dict()
                            self.text_bios.update()
                            self.root.update()  
                            continue
                        # found "*[" but not found value
                        elif re.search(pattern4, line) and (value != tmp_value):
                            self.info_list.insert(self.info_list.len, "Delete *")
                            new_line = line.replace("*[", "[")
                            content[index] = content[index].replace(line, new_line)
                            self.text_bios.see(str(index2)+'.0')
                            self.text_bios.delete(str(index2)+'.0', str(index2)+'.end')
                            self.text_bios.insert(str(index2)+'.0', new_line.strip('\n'))
                            self.make_bios_dict()
                            self.text_bios.update()
                            self.root.update()  
                            continue
                        else:
                            continue
                else:
                    beg2modify = 0
                    beg2modify_opts = 0
            self.list_match.listbox.configure(bg='white')
            self.text_bios.update()
            self.root.update()  
            self.text_bios.config(state='disabled')
            return content
        except Exception as e:
            showerror("Error", str(e))

    # search bios items matched with string in entry box
    def match_items(self, string, data_dict):
        matched_list = []
        pattern = re.compile(r'%s' % string, re.I)
        for key in data_dict.keys():
            if re.search(pattern, key):
                matched_list.append(key)
        self.matched_dataQueue.put(matched_list)
    
    def thread2update(self):
        self.info_list.clear()
        key = self.entry_search.get_string()
        value = self.combobox_value.get_string()
        if key != '' and value != '' and value != self.bios_dict[key]['current']:
            self.info_list.insert(self.info_list.len, "Updating \"%s\", please wait..." % key)
            self.bios_content = self.update_one_bios(self.bios_content, key, value)
            self.info_list.insert(self.info_list.len, "\"%s\" updated to \"%s\"" % (key, value))
        else:
            self.info_list.insert(self.info_list.len, "Nothing changed!")
            
    def callback_update(self):
        threading.Thread(target=self.thread2update).start()

    def callback_open(self):
        fd = askopenfilename(defaultextension='.txt')
        try:
            with open(fd, 'r') as f:
                self.bios_content = f.readlines()
                f.close()
            self.show_bios()
            self.info_list.clear()
            self.info_list.insert(0, "%s has been opened" % fd)
        except Exception as e:
            # ignore no file select exception
            pattern = re.compile(r"No such file or directory", re.I)
            if re.search(pattern, str(e)):
                pass
            else:
                showerror("Exception", str(e))

    def callback_save(self):
        fd = asksaveasfilename(defaultextension='.txt')
        try:
            with open(fd, 'w') as f:
                f.writelines(self.bios_content)
                f.close()
            self.info_list.clear()
            self.info_list.insert(0, "BIOS settings has been saved as %s" % fd)
        except Exception as e:
            # ignore no file select exception
            pattern = re.compile(r"No such file or directory", re.I)
            if re.search(pattern, str(e)):
                pass
            else:
                showerror("Exception", str(e))
    
    def thread2default(self):
        self.info_list.clear()
        self.info_list.insert(self.info_list.len, "Restore BIOS Default...")
        i = 1
        for key in self.bios_dict.keys():
            if self.bios_dict[key]['default'].strip() != 'empty' and self.bios_dict[key]['default'].strip() != self.bios_dict[key]['current'].strip():
                self.bios_content = self.update_one_bios(self.bios_content, key, self.bios_dict[key]['default'])
                self.info_list.insert(self.info_list.len, "\"%s\" has been restored to \"%s\"" % (key, self.bios_dict[key]['default'].strip()))
                self.info_list.see(i)
                i += 1
        self.info_list.insert(self.info_list.len, "BIOS settings restore defaults finished!")
        
    def callback_default(self): 
        threading.Thread(target=self.thread2default).start()
    
    def thread2import(self):
        try: 
            self.info_list.clear()
            self.info_list.insert(0, "Importing BIOS Settings, please wait...")
            with open(bios_import_file, 'w') as fi:
                fi.writelines(self.bios_content)
                fi.close()
            command = "%s /i /s %s" % (sce_tool, bios_import_file)
            with open(import_log, 'w') as fi:
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                process = subprocess.Popen(command, stdout=fi, stderr=fi, stdin=subprocess.PIPE, startupinfo=si)
                process.wait()
                fi.close()
            # show sce import messages to listbox
            self.info_list.clear()
            with open(import_log, 'r') as fi:
                i = 0
                for line in fi.readlines():
                    self.info_list.insert(i, line)
                    i += 1
                fi.close()
            # Reboot to make sense
            if askyesno("Info", "Reboot Now?"):
                os.system("shutdown /r")
        except Exception as e:
                showerror("Exception", str(e))
                
    def callback_import(self):
        threading.Thread(target=self.thread2import).start()
       
    # entry_search keyrelease event
    def event_search_KeyRelease(self, event):
        self.list_match.clear()
        self.match_items(self.entry_search.get_string(), self.bios_dict)
        while True:
            try:
                items = self.matched_dataQueue.get(block=False)
            except queue.Empty:
                break
            else:
                for item in sorted(items):
                    self.list_match.insert(self.list_match.len, item)
                    
    # list_match Button1 Release event
    def event_match_B1Release(self, event):
        self.list_match.get_selection()
        if self.list_match.selection != '':
            self.text_bios.see("%s.0" % self.bios_dict[self.list_match.selection]['index'])
            self.entry_search.set_string(self.list_match.selection)
            self.combobox_value.set_string(self.bios_dict[self.list_match.selection]['current'])
            if 'options' in self.bios_dict[self.list_match.selection].keys():
                self.combobox_value.set_list(self.bios_dict[self.list_match.selection]['options'])
            elif 'value' in self.bios_dict[self.list_match.selection].keys():
                self.combobox_value.set_list(self.bios_dict[self.list_match.selection]['value'])
            else:
                showerror("Error", self.bios_dict[self.list_match.selection])
                
    def event_open(self, event):
        self.callback_open()
        
    def event_save(self, event):
        self.callback_save()
        
    def event_import(self, event):
        self.callback_import()
        
    def event_default(self, event):
        self.callback_default()
    
if __name__ == "__main__":    
    gui = GUIMainWin()
