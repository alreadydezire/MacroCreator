#Runs Task list
import tkinter as tk
from tkinter import ttk
import time



class TaskRunner():
    def __init__(self, master, tasklist, settings, parameters=None):
        print('TaskRunner')
        self.master = master
        self.tasklist = tasklist
        self.settings = settings
        self.parameters = parameters

        try:
            int(self.settings['loopcount'])
            for i in range(0,int(self.settings['loopcount'])):
                self.do_task_loop()
        except:
            paramList=[]
            for param in self.parameters:
                print(param)
                paramList.append(self.getParamsFromFile(param))

        self.do_task_loop()
        self.master.runmacro()

    def getParamsFromFile(self, filename):
        output=[]
        with open(filename,'r') as file:
            for line in file:
                output.append(line.strip())
        return output

    def do_task_loop(self):
        for task in self.tasklist:
            task=task.split(']>  ')[1]
            task=task.split('|')

            match task[0]:
                case "Click":
                    print('click')
                case "String":
                    print('Stringgg')
                case "Keypress":
                    print('Keypress')
                case "Hotkey":
                    print('Hotkey')
                case "Condition":
                    print('Condition')
                case "Wait":
                    print('Wait')
                    time.sleep(int(task[1]))












































