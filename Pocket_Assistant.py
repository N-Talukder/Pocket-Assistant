# importing required packages
from tkinter import *
import tkinter.font as tkFont
from datetime import datetime
import calendar
from tkcalendar import *
import _datetime
import openpyxl
import os
import pandas as pd
from openpyxl.chart import BarChart, Reference
import babel.numbers
from tkinter import ttk




#import Taking_Tag_Choice
class Parent_Program:
    def __init__(self, root, file_Name):
        #global root
        # creating a window and giving a short name 'TK'
        #root = Tk()
        self.root = root
        self.file_Name = file_Name

        # setting the dimension, changing the title and icon
        self.root.geometry('1500x820')
        self.root.title('Pocket Assistant')



        # Create A Main Frame
        self.main_frame = Frame(self.root)
        self.main_frame.pack(fill=BOTH, expand=1)
        # Create A Canvas
        self.my_canvas = Canvas(self.main_frame)
        self.my_canvas.pack(side=LEFT, fill=BOTH, expand=1)
        # Add A Scrollbar To The Canvas
        self.my_vertical_scrollbar = ttk.Scrollbar(self.main_frame, orient=VERTICAL, command=self.my_canvas.yview)
        self.my_vertical_scrollbar.pack(side=RIGHT, fill=Y)
        # Configure The Canvas
        self.my_canvas.configure(yscrollcommand=self.my_vertical_scrollbar.set)
        self.my_canvas.bind('<Configure>', lambda e: self.my_canvas.configure(scrollregion = self.my_canvas.bbox("all")))

        # Add A Scrollbar To The Canvas
        self.my_horizontal_scrollbar = ttk.Scrollbar(self.main_frame, orient=HORIZONTAL, command=self.my_canvas.xview)
        self.my_horizontal_scrollbar.pack(side=BOTTOM, fill=X)
        # Configure The Canvas
        self.my_canvas.configure(xscrollcommand=self.my_horizontal_scrollbar.set)
        self.my_canvas.bind('<Configure>', lambda e: self.my_canvas.configure(scrollregion = self.my_canvas.bbox("all")))


        # Create ANOTHER Frame INSIDE the Canvas
        self.root = Frame(self.my_canvas)
        # Add that New frame To a Window In The Canvas
        self.my_canvas.create_window((0,0), window=self.root, anchor="nw")






        # font specification for the entire display content for user input taking
        self.fontStyle = tkFont.Font(family="Comic Sans MS", size=11, weight = "bold")

        # creating an excel file if there's not any
        # this also asks some information about the user's interest on the accounts to track
        if os.path.isfile(file_Name)  == False:
            from New_Excel_File_Creation import New_Excel_File
            New_Excel_File_Data_Input_Window = New_Excel_File(self.root)
            New_Excel_File_Data_Input_Window.start()

        else:
            self.ModificationFrame = LabelFrame(self.root, borderwidth = 5, padx = 10, pady = 5, width = 1448, height = 90, bg = 'white')
            self.ModificationFrame.grid(row = 0, column = 0, padx = 25, pady = 5, sticky = N, columnspan = 9, ipady = 4)
            self.ModificationFrame.grid_propagate(False)


            self.TransactionFrame = LabelFrame(self.root, borderwidth = 5, padx = 10, pady = 15, width = 640, height = 700, bg = 'white')
            self.TransactionFrame.grid(row = 1, column = 0, padx = 25, pady = 5, sticky = S, columnspan = 3)
            self.TransactionFrame.grid_propagate(False)

            self.StatisticsFrame = LabelFrame(self.root, borderwidth = 5, padx = 10, pady = 15, width = 770, height = 700, bg = 'white')
            self.StatisticsFrame.grid(row = 1, column = 3, padx = 15, pady = 5, sticky = S, columnspan = 3)
            self.StatisticsFrame.grid_propagate(False)

            from Modify_Excel_File import Excel_File_Modification
            Modification_Entry_Window = Excel_File_Modification(root = self.ModificationFrame)

            from Transaction_Entry import Transaction_Entry
            Transaction_Entry_Window = Transaction_Entry(root = self.TransactionFrame)

            from Show_Statistics import Show_Statistics
            Show_Statistics_Window = Show_Statistics(root = self.StatisticsFrame)

    def start(self):
        self.root.mainloop()

if __name__ == '__main__':
    root = Tk()
    Window = Parent_Program(root = root, file_Name = "Balance Sheet.xlsx")
    Window.start()
