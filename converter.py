# imported packages to create this script/program
import os
import pandas as pd
from tkinter import *
from tkinter import messagebox
from docx2pdf import convert

class Window(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.init_window()

    # Creation of init_window
    def init_window(self):

        # changing the title of our master widget
        self.master.title("Word (.docx) To PDF Converter")

        # allowing the widget to take the full space of the root window
        self.pack(fill=BOTH, expand=1)

        # label describing the Run button
        runLabel = Label(self, text="Click 'Run' to start")
        runLabel.pack()
        runLabel.place(x=100, y=125)

        # the Run button and it's attributes
        runButton = Button(self, text="Run", command=self.conversion)
        runButton.place(x=50, y=125)

        # label describing the Run button
        quitLabel = Label(self, text="Click 'Quit' to close")
        quitLabel.pack()
        quitLabel.place(x=100, y=150)

        # the Quit button and it's attributes
        quitButton = Button(self, text="Quit", command=self.client_exit)
        quitButton.place(x=50, y=150)

    # function to find and convert '.docx' files into PDFs
    def conversion(self):

        # indicates current working directory (where this script is located)
        dir = '.'

        # convert all the '.docx' files in this folder location to PDFs
        convert(dir)

        # pop-up message indicating the conversion process is done
        messagebox.showinfo("Good News", "All good to go, you can exit now.")

    # close the program function
    def client_exit(self):
        exit()

root = Tk()

# size of the window
root.geometry("300x300")

app = Window(root)
root.mainloop()