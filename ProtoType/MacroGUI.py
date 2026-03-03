#Add email to SIMs account

from threading import Timer
import time
from msvcrt import getch

#from importstudents import importdata
from ToolTip import CreateToolTip
from AppMenu import menu
from taskgui import taskgui
from draggablelistbox import DragDropListbox
from taskrunner import TaskRunner

import tkinter as tk
from tkinter import ttk
from tkinter import StringVar
from PIL import Image, ImageTk
import pyautogui as p


class macro:
    def __init__(self):
        #Create Window

        self.RUNNING=False

        BORDER_WIDTH = 4
        BORDER_COLOR = 'black'

        WIDGET_COLOR = 'white'#Default:#f0f0f0
        WIDGET_FG = 'black'
        
        TITLE_FONT = ('Yu Gothic UI Semibold',16)
        SUBTITLE_FONT = ('Yu Gothic UI Semibod',12)

        self.filepath = None
        self.filename = 'New Task'
        self.previous_selected = 0
        self.runSettings = {'loopcount':1}
        self.parameterLinks = {}
        '''StudentName':'FILE:names.txt','StudentEmail':'FILE:emails.txt'''
        self.lastOpenWindow=None
        
        self.root=tk.Tk()
        self.root.withdraw()
        self.root.geometry('550x400')
        self.root.title('MacroGUI - '+self.filename)
        self.root.configure(background=BORDER_COLOR)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=4)
        self.root.rowconfigure(1, weight=1)

        self.createImages()#Creates Images

        self.winMenu = menu(self.root, self)
        self.root.config(menu=self.winMenu.menubar)

    #List Frame:
        self.listFrame = tk.Frame(self.root, bg=WIDGET_COLOR)
        self.listFrame.columnconfigure(0, weight=1)
        self.listFrame.rowconfigure(0, weight=1)
        self.listFrame.rowconfigure(1, weight=1)
        self.listFrame.rowconfigure(2, weight=10)
        self.listFrame.grid(row=0,column=0, sticky='nsew',
                            pady=(BORDER_WIDTH,(BORDER_WIDTH/2)),
                            padx=BORDER_WIDTH)
        #List Frame Contents:
            #Title
        self.titleFrame = tk.Frame(self.listFrame, bg=BORDER_COLOR)
        self.titleFrame.columnconfigure(0, weight=1)
        self.titleFrame.rowconfigure(0, weight=1)
        self.titleFrame.grid(row=0,column=0, columnspan=2, sticky='nsew',
                            pady=(0,BORDER_WIDTH))
        titleLabel = tk.Label(self.titleFrame, text='Macro Tasks:',
                              font=(TITLE_FONT), bg=WIDGET_COLOR,
                              fg=WIDGET_FG).grid(column=0, row=0,sticky='nsew',pady=(0,BORDER_WIDTH),ipady=0)
        self.state = tk.StringVar()
        self.state.set('State: Off')
        self.stateLabel = tk.Label(self.titleFrame, textvariable=self.state,
                                   bg=WIDGET_COLOR,fg=WIDGET_FG,anchor='w'
                                   ).grid(column=0, row=1,sticky='nsew',
                                          pady=(0,BORDER_WIDTH),ipady=0)

            #Listbox
        #self.taskListbox = tk.Listbox(self.listFrame, bg=WIDGET_COLOR, selectmode=tk.BROWSE)
        self.taskListbox = DragDropListbox(self.listFrame, bg=WIDGET_COLOR)
        self.listboxVSB = ttk.Scrollbar(self.listFrame, orient=tk.VERTICAL, command=self.taskListbox.yview)
        self.taskListbox.configure(yscroll=self.listboxVSB.set)
        
        self.taskListbox.grid(row=2,column=0,sticky='nsew')
        self.listboxVSB.grid(row=2,column=1, sticky='ns')

            #Listbox BIND EVENTS
        #self.taskListbox.bind('<FocusOut>', lambda e: self.taskListbox.selection_clear(0,'end'))
        self.root.bind('<ButtonPress-1>', self.deselect_item)
        

    #Option Frame:
        self.borderFrame = tk.Frame(self.root,bg='white')
        self.borderFrame.columnconfigure(0, weight=1)
        self.borderFrame.rowconfigure(0, weight=1)
        self.borderFrame.grid(row=1,column=0, sticky='nsew',
                           pady=((BORDER_WIDTH/2),BORDER_WIDTH),
                           padx=BORDER_WIDTH)
        
        self.optFrame = tk.Frame(self.borderFrame,bg='black')
        self.optFrame.columnconfigure(0, weight=1)
        self.optFrame.rowconfigure(0, weight=1)
        self.optFrame.rowconfigure(0, weight=4)
        self.optFrame.rowconfigure(1, weight=2)
        
        self.optFrame.grid(row=0,column=0, sticky='nsew',pady=BORDER_WIDTH,padx=BORDER_WIDTH)
        #Option Frame Contents:
        self.optLabel = tk.Label(self.optFrame, text='Create a task:',bg=WIDGET_COLOR,
                                 fg=WIDGET_FG,font=SUBTITLE_FONT,anchor='w'
                                 ).grid(row=0,column=0, sticky='nsew',pady=(1,0),padx=1)
        self.btnFrame = tk.Frame(self.optFrame,bg=WIDGET_COLOR)
        self.btnFrame.grid(row=1,column=0, sticky='nsew',padx=1,pady=(0,1))
        self.opt_list = [['Click',None,'Create a Click Task.\n    Perform mouse clicks anywhere on screen'],['String',None,'Create a String Task.\n    Types a string using keyboard inputs.'],
                    ['Keypress',None,'Create a Keypress Task.\n    Performs a keypress of any special key (e.g. CTRL, ALT etc)'],['Hotkey',None,'Create a Hotkey Task.\n    Performs any key combination such as; CTRL + ALT + DEL'],
                    ['Condition',None,'Create a Condtion Task.\n    Set a condition that must be satisfied to proceed. (e.g. Pos:x,y must be red)'],['Wait',None,'Create a Wait Task.\n    Wait for specified time to pass until proceeding.'],
                         ['Variable',None,'Create a Variable Task.\n    Creates a variable output that can loop through a list. (e.g. Each loop of macro, type next name in list.)']]
        
        for btn, i in zip(self.opt_list, range(len(self.opt_list))):
            self.btnFrame.columnconfigure(i, weight=1)
            self.btnFrame.rowconfigure(i, weight=1)
            btn[1] = tk.Button(self.btnFrame, text=btn[0], image=self.img_dict[btn[0]],
                               compound='left',bg=WIDGET_COLOR,fg=WIDGET_FG,relief='ridge',
                               command=lambda t=btn[0]:self.createTask_gui(t))
            x_pad=(0,0)
            if i==0:
                x_pad=(4,0)
            elif i==len(self.opt_list)-1:
                x_pad=(0,4)
                
            btn[1].grid(row=0,column=i,sticky='nsew',rowspan=2, padx=x_pad)

            CreateToolTip(btn[1],btn[2])
            
            

        

        self.runBTN = tk.Button(self.optFrame, text='Run', image=self.play_ico, compound='left',
                                command=self.runmacro,bg='lime green', fg='white', relief='ridge',
                                activebackground='forest green',activeforeground='black'
                                )
        self.runBTN.grid(row=2,column=0, sticky='nsew',padx=1,pady=1)

        CreateToolTip(self.runBTN,'Run Macro (Ctrl + Shift)')
        self.update()

        #Hotkey Control
        self.root.bind('<Control-o>',self.winMenu.open)
        self.root.bind('<Control-n>',self.winMenu.new)
        self.root.bind('<Control-s>',self.winMenu.save)
        self.root.bind('<Control-S>',self.winMenu.saveas)

        self.root.bind('<Control-r>',self.runmacro)
        
        self.root.attributes('-topmost',True)

        self.center_window(self.root)
        self.root.update()
        self.root.mainloop()

        

#Update funtion - Runs every second
    def update(self):
        #print('Running:',self.RUNNING)
        #If macro running
        if self.RUNNING:
            pass
        elif not self.RUNNING:
            #Check if tasks are empty
            tasks = list(self.taskListbox.get(0,'end'))
            noTasks_msg="There are currently no tasks, use the options below to start now."
            if len(tasks)==0:
                self.taskListbox.insert(0,noTasks_msg)
                tasks.append(noTasks_msg)
            elif len(tasks)>1 and noTasks_msg in tasks:
                self.taskListbox.delete(tasks.index(noTasks_msg))

            #Disable edit if no item selected
            if self.taskListbox.curselection():
                self.winMenu.menubar.entryconfig("Edit", state='normal')
            else:
                self.winMenu.menubar.entryconfig("Edit", state='disabled')

            #Update window name
            if self.filepath == None:
                self.filename='Untitled Task'
            else:
                self.filename=self.get_name_of_file(self.filepath)
            self.root.title('MacroGUI - '+self.filename)




        
                
        self.root.after(1000,self.update)

#Runs Macro
    def runmacro(self, event=None):
        self.RUNNING = not self.RUNNING #Flip
        if self.RUNNING:
            self.taskListbox.configure(state='disabled')#Disable listbox
            self.state.set('State: Running')


            self.runBTN.configure(text='Stop',image=self.pause_ico, compound='left',
                                command=self.runmacro,bg='red', fg='white', relief='ridge',
                                activebackground='dark red',activeforeground='black')

            ##RUN MACRO:
            TASKLIST = self.taskListbox.get(0,'end')
            SETTINGS = self.runSettings
            PARAMS = self.parameterLinks
            macro = TaskRunner(self,TASKLIST,SETTINGS,PARAMS)
            

            

            
        elif not self.RUNNING:
            self.taskListbox.config(state='normal')#Enable listbox
            self.state.set('State: Terminated')

            self.runBTN.configure(text='Run',image=self.play_ico, compound='left',
                                command=self.runmacro,bg='lime green', fg='white', relief='ridge',
                                activebackground='forest green',activeforeground='black')

    def get_listbox_instructions(self):
        taskItems = self.taskListbox.get(0,'end')
        print(taskItems)
        #tasks = []
        #return tasks
    

    def deselect_item(self,event):
        if self.taskListbox.curselection() == self.previous_selected:
            self.taskListbox.selection_clear(0,'end')
        self.previous_selected = self.taskListbox.curselection()


    def copy_task(self):
        print('Copy Task')

    def delete_task(self):
        self.taskListbox.delete(self.taskListbox.curselection())
        self.taskListbox.create_indexes()


    def createTask_gui(self, taskType):
        #Try destroy last window open
        if self.lastOpenWindow != None:
            try:self.lastOpenWindow.destroy()
            except:pass
        taskgui(self, 'new', taskType)

    def get_name_of_file(self,filepath):
        i = len(filepath)-1
        while True:
            if filepath[i]=='/':
                filename = filepath[i+1:]
                break
            else:
                i-=1
        filename=filename[:len(filename)-4]
        return filename

    def center_window(self,root):
        root.update()

        app_width=root.winfo_width()
        app_height=root.winfo_height()

        x=(root.winfo_screenwidth()-app_width)/2
        y=(root.winfo_screenheight()-app_height)/2

        root.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')
        root.deiconify()
    

    def createImages(self):
        self.img_dict={}
        #Settings Icon
        self.image = Image.open('Images/sett-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.settings_ico= ImageTk.PhotoImage(self.image)

        #Play Icon
        self.image = Image.open('Images/play-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.play_ico= ImageTk.PhotoImage(self.image)

        #Pause Icon
        self.image = Image.open('Images/pause-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.pause_ico= ImageTk.PhotoImage(self.image)

        #Click Icon
        self.image = Image.open('Images/Tasks/click-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.click_ico= ImageTk.PhotoImage(self.image)
        self.img_dict['Click']=self.click_ico

        #Condition Icon
        self.image = Image.open('Images/Tasks/condition-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.condition_ico= ImageTk.PhotoImage(self.image)
        self.img_dict['Condition']=self.condition_ico

        #Press
        self.image = Image.open('Images/Tasks/press-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.press_ico= ImageTk.PhotoImage(self.image)
        self.img_dict['Keypress']=self.press_ico

        #Type
        self.image = Image.open('Images/Tasks/type-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.type_ico= ImageTk.PhotoImage(self.image)
        self.img_dict['String']=self.type_ico

        #Hotkey
        self.image = Image.open('Images/Tasks/hotkey-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.hotkey_ico= ImageTk.PhotoImage(self.image)
        self.img_dict['Hotkey']=self.hotkey_ico

        #Wait
        self.image = Image.open('Images/Tasks/wait-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.wait_ico= ImageTk.PhotoImage(self.image)
        self.img_dict['Wait']=self.wait_ico

        #Variable Icon
        self.image = Image.open('Images/Tasks/variable-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.variable_ico= ImageTk.PhotoImage(self.image)
        self.img_dict['Variable']=self.variable_ico

        #Position Icon
        self.image = Image.open('Images/position.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.position_ico= ImageTk.PhotoImage(self.image)

        #Dropper Icon
        self.image = Image.open('Images/dropper.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.dropper_ico= ImageTk.PhotoImage(self.image)

        #Find Keypress Icon
        self.image = Image.open('Images/find-keypress.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.findkey_ico= ImageTk.PhotoImage(self.image)





app = macro()
