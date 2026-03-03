import tkinter as Tkinter



class DragDropListbox(Tkinter.Listbox):
    """ A Tkinter listbox with drag'n'drop reordering of entries. """
    def __init__(self, master, **kw):
        kw['selectmode'] = Tkinter.SINGLE
        Tkinter.Listbox.__init__(self, master, kw)
        self.bind('<Button-1>', self.setCurrent)
        self.bind('<B1-Motion>', self.shiftSelection)
        self.curIndex = None

    def setCurrent(self, event):
        self.curIndex = self.nearest(event.y)
       
    def shiftSelection(self, event):
        i = self.nearest(event.y)
        if i < self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i+1, x)
            self.curIndex = i
        elif i > self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i-1, x)
            self.curIndex = i
        self.create_indexes()

    def create_indexes(self):
        noTasks_msg="There are currently no tasks, use the options below to start now."
        for item in self.get(0,'end'):
            try:
                half = item.split(']>  ')[1]
            except:
                half = item
            if noTasks_msg not in item:
                try:
                    newdata = "  ["+str(self.get(0,'end').index(item)) + "]>  " +half
                    edit_item = self.get(0,'end').index(item)
                    self.delete(edit_item)
                    self.insert(edit_item, newdata)
                except:pass















































