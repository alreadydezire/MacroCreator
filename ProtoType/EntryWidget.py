from tkinter import ttk
import tkinter as tk

import time
import pyautogui as p
import keyboard
from ToolTip import CreateToolTip
from functools import partial
from PIL import ImageGrab

class FormEntry(ttk.Frame):
    def __init__(self, master, header, widget,*args, **kw):
        super().__init__(master.form_frame, **kw)

        self.master = master.root
        self.macro = master.master
        self.parent=master
        self.header = header
        self.widget = widget

        try:
            self.wid_type = self.widget.split(':')[0]
            self.wid_data = self.widget.split(':')[1]
        except:self.wid_type=self.widget

        keys = ['!', '"', '#', '$', '%', '&', "'", '(',')', '*', '+', ',', '-', '.', '/', '0', '1', '2', '3', '4', '5', '6', '7',
                     '8', '9', ':', ';', '<', '=', '>', '?', '@', '[', '\\', ']', '^', '_', '`','a', 'b', 'c', 'd', 'e','f', 'g', 'h',
                     'i', 'j', 'k', 'l', 'm', 'n', 'o','p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '{', '|', '}', '~','accept',
                     'add', 'alt', 'altleft', 'altright', 'apps', 'backspace','browserback', 'browserfavorites', 'browserforward',
                     'browserhome','browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear','convert', 'ctrl', 'ctrlleft',
                     'ctrlright', 'decimal', 'del', 'delete','divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10',
                     'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20','f21', 'f22', 'f23', 'f24', 'f3', 'f4',
                     'f5', 'f6', 'f7', 'f8', 'f9','final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja','kana',
                     'kanji', 'launchapp1', 'launchapp2', 'launchmail','launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack',
                     'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6','num7', 'num8', 'num9', 'numlock', 'pagedown',
                     'pageup', 'pause', 'pgdn','pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn','prtsc', 'prtscr',
                     'return', 'right', 'scrolllock', 'select', 'separator','shift', 'shiftleft', 'shiftright', 'sleep', 'space', 'stop',
                     'subtract', 'tab','up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen','command', 'option',
                     'optionleft', 'optionright']
        
        # Types of dropdowns
        self.drop_options={'clicks':['Single','Double'],
                           'keys':keys,
                           'hkeys':[['Ctrl','Ctrl+Shift','Ctrl+Alt'],['a', 'b', 'c', 'd', 'e','f', 'g', 'h','i', 'j', 'k', 'l', 'm', 'n', 'o','p',
                                                                      'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']],
                           'mkeys':['Left', 'Right', 'Middle'],
                           'params':list(self.macro.parameterLinks.keys())
                           }

        self.columnconfigure(0,weight=1)
        self.columnconfigure(1,weight=3)


        self.entry_Label = ttk.Label(self, text=self.header,width=15).grid(row=0,column=0,sticky='w',padx=(0,15))

        # XY Position Entry
        if self.wid_type == 'posent':
            self.entry = ttk.Entry(self,width=13)
            self.entry_button = ttk.Button(self, command=self.getpos, width=5,image=self.macro.position_ico)
            self.entry_button.grid(row=0,column=2,sticky='e')
            CreateToolTip(self.entry_button, "Get Position.\n    Press 'CTRL' to get x,y position from mouse location.\n    Press 'ESC' to cancel.")

        # Dropdown Entry
        elif self.wid_type=='drop':
            #If entry type is Hot Keys
            if self.wid_data == 'hkeys':
                #Create dropdown for each key
                self.entry = []
                for i in range(0,2):
                    key_frame = tk.Frame(self, borderwidth=4, relief = 'groove')
                    variable = tk.StringVar(key_frame)
                    dropdown = ttk.Combobox(key_frame, textvariable=variable, values=self.drop_options[self.wid_data][i],state='readonly',width=9)
                    dropdown.grid(row=0,column=0)
                    key_button = ttk.Button(key_frame, command=lambda var=variable:self.getkey(var), image = self.macro.findkey_ico)
                    key_button.grid(row=0,column=1)
                    CreateToolTip(key_button, "Get key from keypress.")
                    self.entry.append(variable)
                    key_frame.grid(row=0,column=i+1)

            #If not hotkey
            else:
                #Create default dropdown
                self.entry_Options = self.drop_options[self.wid_data]
                self.entry = tk.StringVar(self)
                self.entry_drop = ttk.Combobox(self, textvariable=self.entry, values=self.entry_Options,state='readonly',width=15)
                self.entry_drop.grid(row=0,column=1,sticky='e')

                #If type is keypress
                if self.wid_data == 'keys':
                    #Button to get keyname from keypress
                    self.entry_button = ttk.Button(self, command=self.getkey, image = self.macro.findkey_ico)
                    self.entry_button.grid(row=0,column=2)
                    CreateToolTip(self.entry_button, "Get key from keypress.")
                    self.entry.set('[KEY]')
                #if not keypress or hotkey
                else:
                    #Try setting entry to first dropdown option
                    try:self.entry.set(self.entry_Options[0])
                    except:
                        #Disable dropdown
                        self.entry.set('No Sources')
                        self.entry_drop.config(state='disabled')
                        self.parent.confirm_btn.config(state='disabled')
                        CreateToolTip(self.entry_drop, "Variable sources can be set up in the 'Macro' Menu. \n(Macro>Configure Macro>Manage Parameters)")
                        CreateToolTip(self.parent.confirm_btn, "Variable sources can be set up in the 'Macro' Menu. \n(Macro>Configure Macro>Manage Parameters)")

        # RGB Color Entry
        elif self.wid_type == 'rgbent':
            self.entry = ttk.Entry(self,width=13)
            self.entry_button = ttk.Button(self, command=self.getcol,width=5,image=self.macro.dropper_ico)
            self.entry_button.grid(row=0,column=2)
            CreateToolTip(self.entry_button, "Get Color.\n    Press 'CTRL' to get RGB values from position of mouse.\n    Press 'ESC' to cancel.")

        # Text Entry
        elif self.wid_type == 'ent':
            self.entry = ttk.Entry(self)
            
        #If entry widget (Entry):
        else:self.entry = ttk.Entry(self)

        try:self.entry.grid(row=0,column=1, sticky='e',pady=(5,0))   
        except:pass

    def getcol(self):
        ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
        while True:
            key=keyboard.read_key()
            if key=='ctrl':
                img=p.screenshot()
                #print(p.position)
                col = img.getpixel(p.position())
                break
            elif key=='esc':
                break
        self.entry.delete(0,'end')
        self.entry.insert(0, ','.join(str(i) for i in col))

    def getpos(self):
        while True:
            key=keyboard.read_key()
            if key=='ctrl':
                pos = p.position()
                break
            elif key=='esc':
                break
        self.entry.delete(0,'end')
        self.entry.insert(0, ','.join(str(i) for i in pos))
        #self.entry.set(','.join(str(i) for i in pos))

    def getkey(self, entry):
        key=keyboard.read_key()
        entry.set(key)

            



        
