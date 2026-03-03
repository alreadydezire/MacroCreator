#Add email to SIMs account
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk

from tkinter import filedialog as fd
from tkinter import messagebox as mb

#from PIL import Image, ImageTk
files = [('Text Document', '*.txt')]

noTasks_msg="There are currently no tasks, use the options below to start now."

class menu():
    def __init__(self, master,masterclass):

        self.master=master
        self.masterclass=masterclass

        self.create_images()
        
        self.menubar = tk.Menu(self.master)

        #File Menu - (New,Open,Save,Saveas)
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="New File", accelerator='Ctrl+N', command=self.new,image=self.new_ico,compound=tk.LEFT)
        self.filemenu.add_command(label="Open...", accelerator='Ctrl+O', command=self.open,image=self.open_ico,compound=tk.LEFT)
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Save", accelerator='Ctrl+S', command=self.save,image=self.save_ico,compound=tk.LEFT)
        self.filemenu.add_command(label="Save As...", accelerator='Ctrl+Shift+S', command=self.saveas,image=self.saveas_ico,compound=tk.LEFT)
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", accelerator='Ctrl+Q', command=self.masterclass.root.quit(),image=self.exit_ico,compound=tk.LEFT)
        self.menubar.add_cascade(label="File", menu=self.filemenu)

    #Edit Menu - (Copy, Paste, Delete)
        self.editmenu = tk.Menu(self.master, tearoff=0)
        self.editmenu.add_command(label="Copy", accelerator='Ctrl+C', command=lambda e='copy':self.donothing(e),image=self.copy_ico,compound=tk.LEFT)
        self.editmenu.add_command(label="Paste", accelerator='Ctrl+V', command=lambda e='paste':self.donothing(e),image=self.paste_ico,compound=tk.LEFT)
        self.editmenu.add_command(label="Delete", accelerator='Ctrl+Del', command=self.masterclass.delete_task,image=self.delete_ico,compound=tk.LEFT)
        self.menubar.add_cascade(label="Edit", menu=self.editmenu)

    #Macro Menu
        self.macromenu = tk.Menu(self.master, tearoff=0)
        self.macromenu.add_command(label="Run Macro", command=self.masterclass.runmacro,image=self.greenplay_ico,compound=tk.LEFT)
        self.macromenu.add_separator()
         #Create Task Submenu
        self.taskmenu = tk.Menu(self.macromenu, tearoff=0)
        self.taskmenu.add_command(label="Mouse Click", command=lambda t='Click':self.masterclass.createTask_gui(t),image=self.masterclass.img_dict['Click'],compound=tk.LEFT)
        self.taskmenu.add_command(label="Type String", command=lambda t='String':self.masterclass.createTask_gui(t),image=self.masterclass.img_dict['String'],compound=tk.LEFT)
        self.taskmenu.add_command(label="Keypress", command=lambda t='Keypress':self.masterclass.createTask_gui(t),image=self.masterclass.img_dict['Keypress'],compound=tk.LEFT)
        self.taskmenu.add_command(label="Hotkey", command=lambda t='Hotkey':self.masterclass.createTask_gui(t),image=self.masterclass.img_dict['Hotkey'],compound=tk.LEFT)
        self.taskmenu.add_command(label="Condition", command=lambda t='Condition':self.masterclass.createTask_gui(t),image=self.masterclass.img_dict['Condition'],compound=tk.LEFT)
        self.taskmenu.add_command(label="Wait", command=lambda t='Wait':self.masterclass.createTask_gui(t),image=self.masterclass.img_dict['Wait'],compound=tk.LEFT)
        self.taskmenu.add_command(label="Variable", command=lambda t='Variable':self.masterclass.createTask_gui(t),image=self.masterclass.img_dict['Variable'],compound=tk.LEFT)
        self.macromenu.add_cascade(label="Create Task", menu=self.taskmenu,image=self.createtask_ico,compound=tk.LEFT)
         #Configure Macro Submenu
        self.optionsubmenu = tk.Menu(self.macromenu, tearoff=0)
        self.optionsubmenu.add_command(label="Manage Parameters", command=lambda e='Params':self.donothing(e),image=self.varsett_ico,compound=tk.LEFT)
        self.optionsubmenu.add_command(label="Run Settings", command=lambda e='Settings':self.donothing(e),image=self.runsett_ico,compound=tk.LEFT)
        self.macromenu.add_cascade(label="Configure Macro", menu=self.optionsubmenu,image=self.config_ico,compound=tk.LEFT)
        self.menubar.add_cascade(label="Macro", menu=self.macromenu)
     
    #Other Menu
        self.othermenu = tk.Menu(self.menubar, tearoff=0)
        self.othermenu.add_command(label="Help Index", command=lambda e='help':self.donothing(e),image=self.help_ico,compound=tk.LEFT)
        self.othermenu.add_command(label="About...", command=lambda e='about':self.donothing(e),image=self.about_ico,compound=tk.LEFT)
        self.menubar.add_cascade(label="Other", menu=self.othermenu)

    def donothing(self,option=None):
        print(option)


    def new(self, event=None):
        #print('New File')
        #filepath = fd.asksaveasfilename(filetypes=files,defaultextension='.txt')
        #if filepath:

        res = mb.askquestion('New Task File', 'Are you sure you want to create a new task?\nAny unsaved changed will be lost.')
        
        if res == 'yes':
            #reset variables
            self.masterclass.filepath=None

            self.masterclass.taskListbox.delete(0,'end')

            self.masterclass.runSettings={'loopcount':1}

            self.masterclass.parameterLinks = {}
            

        else:
            mb.showinfo('Return','Returning to main application')
        
            #with open(filepath,'a') as file:
                #file.write('')


    def open(self, event=None):
        #print('Open File')
        res = mb.askquestion('Open File', 'Are you sure you want to open a task file?\nAny unsaved changed will be lost.')
        if res == 'yes':
            
            filepath = fd.askopenfilename()
            if filepath:
                self.masterclass.filepath=filepath
                
                with open(filepath,'r') as file:
                    readType=None
                    self.masterclass.taskListbox.delete(0,'end')
                    for line in file:
                        line=line.strip()
                        
                        if line=='<task>' or line=='<sett>' or line=='<param>':
                            readType=line
                            continue
                            
                        if readType=='<task>':
                            self.masterclass.taskListbox.insert('end',line)
                        if readType=='<sett>':
                            self.masterclass.runSettings[line.split('|')[0]]=line.split('|')[1]
                        if readType=='<param>':
                            try:self.masterclass.parameterLinks[line.split('|')[0]]=line.split('|')[1]
                            except:pass
                            
                self.masterclass.taskListbox.create_indexes()



    def saveas(self, event=None):
        #print('Save File')
        filepath = fd.asksaveasfilename(filetypes=files, defaultextension='.txt')
        if filepath:
            self.masterclass.filepath=filepath
            fileContent = self.create_file()
            with open(filepath,'a') as file:
                for line in fileContent:
                    file.write(line+'\n')

    def save(self, event=None):
        if not self.masterclass.filepath:
            self.saveas()
        else:
            fileContent = self.create_file()
            with open(self.masterclass.filepath,'w') as file:
                for line in fileContent:
                    file.write(line+'\n')

    def create_file(self):
        tasklist = self.masterclass.taskListbox.get(0,'end')
        settings = self.masterclass.runSettings
        parameters = self.masterclass.parameterLinks

        file_lines = []
        file_lines.append('<task>')
        for task in tasklist:
            file_lines.append(task.split(']>  ')[1])
        file_lines.append('<sett>')
        for setting in settings:
            file_lines.append(setting+"|"+str(settings[setting]))
        file_lines.append('<param>')
        for parameter in parameters:
            file_lines.append(parameter)

        return file_lines
        

    def create_images(self):
        #New File Icon
        self.image = Image.open('Images/Menu/File/new.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.new_ico= ImageTk.PhotoImage(self.image)
        #Open File Icon
        self.image = Image.open('Images/Menu/File/open.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.open_ico= ImageTk.PhotoImage(self.image)
        #Save Icon
        self.image = Image.open('Images/Menu/File/save.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.save_ico= ImageTk.PhotoImage(self.image)
        #Save as Icon
        self.image = Image.open('Images/Menu/File/saveas.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.saveas_ico= ImageTk.PhotoImage(self.image)
        #Exit Icon
        self.image = Image.open('Images/Menu/File/exit.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.exit_ico= ImageTk.PhotoImage(self.image)
        #Copy Icon
        self.image = Image.open('Images/Menu/Edit/copy.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.copy_ico= ImageTk.PhotoImage(self.image)
        #Paste Icon
        self.image = Image.open('Images/Menu/Edit/paste.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.paste_ico= ImageTk.PhotoImage(self.image)
        #Delete Icon
        self.image = Image.open('Images/Menu/Edit/delete.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.delete_ico= ImageTk.PhotoImage(self.image)
        #Help Icon
        self.image = Image.open('Images/Menu/Other/help.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.help_ico= ImageTk.PhotoImage(self.image)
        #Info Icon
        self.image = Image.open('Images/Menu/Other/about.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.about_ico= ImageTk.PhotoImage(self.image)
        #Green Play Icon
        self.image = Image.open('Images/Menu/Macro/greenplay-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.greenplay_ico= ImageTk.PhotoImage(self.image)
        #Config Icon
        self.image = Image.open('Images/Menu/Macro/config.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.config_ico= ImageTk.PhotoImage(self.image)
        #Create Task Icon
        self.image = Image.open('Images/Menu/Macro/task-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.createtask_ico= ImageTk.PhotoImage(self.image)
        #Run Settings Icon
        self.image = Image.open('Images/Menu/Macro/runsett-ico.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.runsett_ico= ImageTk.PhotoImage(self.image)
        #Var Settings Icon
        self.image = Image.open('Images/Menu/Macro/variable-setting.png')
        self.image = self.image.resize((20,20), Image.LANCZOS)
        self.varsett_ico= ImageTk.PhotoImage(self.image)



