#Root + dynamic Treeview page
import tkinter as tk
from tkinter import ttk

from EntryWidget import FormEntry


class taskgui:
    def __init__(self,master,form_type,ref):
        self.master=master
        self.root = tk.Toplevel()

        #self.root.geometry('500x100')

        self.master.lastOpenWindow=self.root
        
        self.root.withdraw()
        
        self.form_type=form_type
        self.ref=ref

        self.root.title('Create Task - '+self.ref)

        self.root.columnconfigure(0,weight=1)

        self.content_frame = ttk.Frame(self.root)
        self.content_frame.columnconfigure(0,weight=1)
        self.content_frame.rowconfigure(0,weight=1)
        self.content_frame.grid(row=0,column=0, sticky='nsew')
        
        self.form_frame= ttk.Frame(self.content_frame)
        self.form_frame.columnconfigure(0,weight=1)
        self.form_frame.grid(row=0,column=0, padx=10, sticky='nsew')
        self.btn_frame= ttk.Frame(self.content_frame)
        self.btn_frame.grid(row=1,column=0, pady=10)
        self.cancel_btn = ttk.Button(self.btn_frame,text='Cancel',command=self.cancel).grid(row=0,column=0,sticky='ew')
        self.confirm_btn = ttk.Button(self.btn_frame,text='Save Task',command=self.formcomplete)
        self.confirm_btn.grid(row=0,column=1,sticky='ew')

        #Create
        self.form_entries=self.createForm()
        #position_root(self.root)

        self.root.attributes('-topmost',True)

        self.master.center_window(self.root)
        self.root.update()
        
        self.root.mainloop()

    def __del__(self):
        print('im gone')

    #Create form's widgets
    def createForm(self):

        task_form = {'Click':[['Mouse Position:','posent'],['Click Type:','drop:clicks'],['Mouse Button:', 'drop:mkeys']],
                     'String':[['Target String:', 'ent']],
                     'Keypress':[['Target Key:','drop:keys']],
                     'Hotkey':[['Hotkey Combo:','drop:hkeys']],
                     'Condition':[['Mouse Position:','posent'],['Color (RGB):','rgbent']],
                     'Wait':[['Time (Seconds):','ent']],
                     'Variable':[['Variable Source:','drop:params']]}
        headers=[]
        widgets=[]
        for entry in task_form[self.ref]:
            headers.append(entry[0])
            widgets.append(entry[1])
        
        form_entries=[]
        #create widgets
        for header, widget, index in zip(headers, widgets,range(0, len(headers))):
            entry_frame= FormEntry(self,header, widget)
            form_entries.append(entry_frame.entry)
            entry_frame.grid(row=index,column=0,sticky='nsew',padx=30)
        return form_entries

    #On form submit button press
    def formcomplete(self):
        input_arr=[]
        #Get all inputs from Entries
        for widget in self.form_entries:
            try:input_arr.append(widget.get())
            except:
                try:input_arr.append(widget.get(0,'end'))
                except:
                    try:
                        for ent in widget:
                            input_arr.append(ent.get())
                    except:print("ERROR: Can't get entry data")
        if self.ref=='Hotkey':
            self.master.taskListbox.insert('end',self.ref+'|'+'+'.join(input_arr))
        else:
            self.master.taskListbox.insert('end',self.ref+'|'+'|'.join(input_arr))
        self.master.taskListbox.create_indexes()
        
        #self.master.table.refresh()
        self.cancel()


    def cancel(self):
        self.master.root.deiconify()
        self.root.destroy()




#app=taskgui('new','Click')
































































