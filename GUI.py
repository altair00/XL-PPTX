from tkinter import *
from text_top import *
from image_top import *
from tkinter import ttk
import tkinter.font as font
from icon_method import *
import os
import base64


class GUI:
    def __init__(self, master):
        self.master = master
        master.title("A Program")
        master.geometry("500x200")
        gui_font = font.Font(family='Times New Roman', weight='bold', size=20)
        ttk.Style(master).configure("TButton", font=gui_font)

        self.text_top_bt = ttk.Button(master, text="Text", command=self.text_top, style="TButton")
        self.text_top_bt.pack(fill=BOTH, side=LEFT, expand=True, padx=20, pady=20)

        self.image_top_bt = ttk.Button(master, text="Image", command=self.image_top)
        self.image_top_bt.pack(fill=BOTH, side=LEFT, expand=True, padx=20, pady=20)




    def text_top(self):
        self.t_top = Toplevel()
        self.t_top.title("Text Operation")
        self.t_top.geometry("535x100")
        self.t_top.resizable(False, False)
        text_top(self.t_top)
        self.t_top.after(50, self.t_top.iconbitmap(os.path.join(os.getcwd(), 'icon.ico')))

        
        
        


    def image_top(self): 
        self.i_top = Toplevel()
        self.i_top.title("Image Operation")
        self.i_top.geometry("535x100")
        self.i_top.resizable(False, False)
        image_top(self.i_top)
        self.i_top.after(50, self.i_top.iconbitmap(os.path.join(os.getcwd(), 'icon.ico')))





if __name__ == "__main__":
    root = Tk()
    save_ico()
    gui = GUI(root)
    root.resizable(False, False)
    root.after(50, root.iconbitmap(os.path.join(os.getcwd(), 'icon.ico')))
    root.mainloop()
    os.remove('icon.ico')
    