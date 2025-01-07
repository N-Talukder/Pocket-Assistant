from tkinter import *
import tkinter.font as tkFont
import pandas as pd

# funcion for taking tag entries from user
class Tag_Options_User_Choice:
    def __init__(self, root, Name_Entry, writer):
        self.root = root
        self.Name_Entry = Name_Entry
        self.writer = writer

        self.fontStyle = tkFont.Font(family="Comic Sans MS", size=11, weight = "bold")

        # Frame for taking user's choice for dividing transactions according to his/her choice into different categories
        self.Tag_Options_Frame = LabelFrame(root, text = "User's Choice for Transaction Tags", borderwidth = 5, padx = 10, pady = 15, width = 1430, height = 700, bg = 'white', font = self.fontStyle)
        self.Tag_Options_Frame.grid(row = 0, column = 0, padx = 35, pady = 10)
        #self.Tag_Options_Frame.grid_propagate(False)

        self.User_Choice_Tag_Text = Label(self.Tag_Options_Frame, text = "Click 'Add' for the tags preset as your choice for tracking transactions and when not found, you can enter them in the emtpy text boxes below:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 0, column = 0, columnspan = 12, pady = 10, sticky = W)

        # function that will enter the text left to the 'Add' button into the empty box to the right
        def Add_Tag(x):
            self.Button_Entry_Dictionary["Tag_User_Choice_Entry" + str(x)].insert(0, self.Tag_Options_Preset[x] )
            return

        # creating preset tag choices, add buttons, empty boxes for user inputs of tags to classify data later using those tags
        self.Tag_Options_Preset = ["Betel Leave", "Bill", "Donation", "Fish", "Fruit", "Home Maintenance", "Avoidable Grocery", "Grocery", "Medicine", "Milk", "Mobile", "Money", "Society Incurred Cost", "System Loss", "Tea", "Transportation", "Vegetables", Name_Entry.get() + "'s Personal Expense"]

        t = 0
        r = 0
        c = 0
        global Button_Entry_Dictionary
        self.Button_Entry_Dictionary = {}

        for x in self.Tag_Options_Preset:
            # displaying the preset tags as texts onto the frame
            self.currentPreset = Label(self.Tag_Options_Frame, text = x, justify = LEFT, bg = 'white', font = self.fontStyle)
            self.currentPreset.grid(row = r + 1, column = c, sticky = E, padx = 10)

            # creating add buttons to the right of those text preset tags
            self.Button_Entry_Dictionary["addButton" + str(t)] = Button(self.Tag_Options_Frame, text = "Add", width = 21, font = self.fontStyle, background = "grey", foreground = "white", command = lambda t=t: Add_Tag(t))
            self.Button_Entry_Dictionary["addButton" + str(t)].grid(row = r + 1, column = c + 1, padx = 5, pady = 5)

            # creating empty boxes to enter the text left to the 'Add' button to the boxes when clicked
            self.Button_Entry_Dictionary["Tag_User_Choice_Entry" + str(t)] = Entry(self.Tag_Options_Frame, width = 21,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
            self.Button_Entry_Dictionary["Tag_User_Choice_Entry" + str(t)].grid(row = r + 1, column = c + 2, columnspan = 1, sticky = W, padx = 10, pady = 5)

            # controlling the appearance on the frame using r for row and c for column and t controls the dictionary entries' numbers for referring to them later
            if r > 7:
                r = 0
                c = 3
            else:
                r += 1
            t += 1

        # Creating blank boxes for user inputs of tags to be used in data classification and storing them in the dictionary previously created for preset entries
        for x in range(6):
            x = x + 18
            self.Button_Entry_Dictionary["Tag_User_Choice_Entry" + str(x)] = Entry(self.Tag_Options_Frame, width = 21,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
            self.Button_Entry_Dictionary["Tag_User_Choice_Entry" + str(x)].grid(row = 10, column = x - 18, pady = 25, padx = 10)

        # Creating Submit button that will store the user preferences of tags and will take to the transaction entry and statistics related frame
        self.SubmitButton = Button(self.Tag_Options_Frame, text = "Submit", width = 30, background = "green", foreground = "white", font = self.fontStyle, command = self.Final_Tag_Entry).grid(row = 11, column = 1, columnspan = 7, padx = 5, pady = 10)


    def Final_Tag_Entry(self):

        # Creating a list of all text boxes in the tag entry frame
        #global Tags
        self.Tags = []
        for j in range(24):
            self.Tags.append(self.Button_Entry_Dictionary["Tag_User_Choice_Entry" + str(j)].get())
        # getting rid of the empty text boxes or the tags discarded by the user
        self.Tags = [Tag for Tag in self.Tags if Tag != ""]

        self.Tags_Data_Frame = pd.DataFrame(self.Tags, index = None, columns = None)
        # writing the tags to the excel file so that those tags are also accessible when this function won't be called once the user input tags are entered
        #global writer
        self.Tags_Data_Frame.to_excel(self.writer, sheet_name = "Tags", index = None, startrow = 0, startcol = 0, header = None)
        # saving the tags in the excel file
        self.writer.close()#save()
        # newly created frames for transaction entry and statistics will be over
        # this tags frame if not destroyed. So, deleting those widgets for making
        # sure that those tag entry widgets are not visible from back
        for widgets in self.root.winfo_children():
            widgets.destroy()
        # once hte excel file is created it's time to display options for entering transactions, which is done by calling function Edit_Excel_File
        self.ModificationFrame = LabelFrame(self.root, borderwidth = 5, padx = 10, pady = 5, width = 1448, height = 90, bg = 'white')
        self.ModificationFrame.grid(row = 0, column = 0, padx = 25, pady = 5, sticky = N, columnspan = 7, ipady = 4)
        #self.ModificationFrame.grid_propagate(False)

        self.TransactionFrame = LabelFrame(self.root, borderwidth = 5, padx = 10, pady = 15, width = 640, height = 700, bg = 'white')
        self.TransactionFrame.grid(row = 1, column = 0, padx = 25, pady = 5, sticky = S, columnspan = 3)
        #self.TransactionFrame.grid_propagate(False)

        self.StatisticsFrame = LabelFrame(self.root, borderwidth = 5, padx = 10, pady = 15, width = 770, height = 700, bg = 'white')
        self.StatisticsFrame.grid(row = 1, column = 3, padx = 15, pady = 5, sticky = S, columnspan = 3)
        #self.StatisticsFrame.grid_propagate(False)

        from Modify_Excel_File import Excel_File_Modification
        Modification_Entry_Window = Excel_File_Modification(root = self.ModificationFrame)

        from Transaction_Entry import Transaction_Entry
        Transaction_Entry_Window = Transaction_Entry(root = self.TransactionFrame)

        from Show_Statistics import Show_Statistics
        Show_Statistics_Window = Show_Statistics(root = self.StatisticsFrame)


    def start(self):
        self.root.mainloop()
