from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from text_functions import *
import os


class text_top:
    def __init__(self, top):
        Label(top, text="Select an excel file: ").grid(row=0, padx=10)
        Label(top, text="Select a ppt file: ").grid(row=1, padx=10)
        
        self.ppt_entry_box = Entry(top, width=50)
        self.ppt_entry_box.grid(row=1, column=1)

        self.xl_entry_box = Entry(top, width=50)
        self.xl_entry_box.grid(row=0, column=1)

        self.xl_entry_btn = Button(top, text="Open xl Files", command= lambda: self.open(self.xl_entry_box))
        self.xl_entry_btn.grid(row=0, column=2, sticky='news', padx=10)

        self.ppt_entry_btn = Button(top, text="Open ppt Files", command=lambda: self.open(self.ppt_entry_box))
        self.ppt_entry_btn.grid(row=1, column=2, sticky='news', padx=10)

        self.execute_btn = Button(top, text="Execute", command=self.execute)
        self.execute_btn.grid(row=3, column=1)
    
    def execute(self):
        if(".xlsx" in self.xl_entry_box.get() and ".pptx" in self.ppt_entry_box.get()):
            t_runner(self.xl_entry_box.get(), self.ppt_entry_box.get())
        else:
            messagebox.showerror(title="Error", message="Check the imported file again")

    def open(self, box):
        if(box==self.xl_entry_box):
            filename = filedialog.askopenfilename(title="Select A File", filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if(box==self.ppt_entry_box):
            filename = filedialog.askopenfilename(title="Select A File", filetypes=(("power poin files", "*.pptx"), ("all files", "*.*")))
        box.delete(0, END)
        box.insert(0, filename)