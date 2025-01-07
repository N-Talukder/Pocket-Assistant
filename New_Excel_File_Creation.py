from tkinter import *
import tkinter.font as tkFont
import pandas as pd
from datetime import datetime
import _datetime


class New_Excel_File:
    def __init__(self, root):
        self.root = root
        self.fontStyle = tkFont.Font(family="Comic Sans MS", size=11, weight = "bold")

        # Parent Frame : creating a parent frame for user and account information inputs
        self.Information_Frame = LabelFrame(self.root, borderwidth = 5, padx = 10, pady = 15, width = 1430, height = 700, bg = 'white', font = self.fontStyle)
        self.Information_Frame.grid(row = 0, column = 0, padx = 35, pady = 10)
        #Information_Frame.#grid_propagate(False)


        # Child Frame 1 :  User Information Frame
        self.User_Information_Frame = LabelFrame(self.Information_Frame, text = "User Information", borderwidth = 5, padx = 10, pady = 15, width = 1400, height = 150, bg = 'white', font = self.fontStyle)
        self.User_Information_Frame.grid(row = 0, column = 0, sticky = W, pady = 0, columnspan = 2)
        #self.User_Information_Frame.#grid_propagate(False)

        self.NameText = Label(self.User_Information_Frame, text = "                   Name:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 0, column = 0, sticky = E, rowspan = 2)
        global Name_Entry
        self.Name_Entry = Entry(self.User_Information_Frame, width = 34,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
        self.Name_Entry.grid(row = 0, column = 1, columnspan = 2, sticky = W, padx = 10, pady = 5, rowspan = 2)

        self.BankAccountNoText = Label(self.User_Information_Frame, text = "Number of Bank Accounts:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 0, column = 3, sticky = E)
        global No_Bank_Account
        self.No_Bank_Account = StringVar(self.root)
        self.No_Bank_Account.set("")
        self.No_Bank_Account_DropDown_Menu = OptionMenu(self.User_Information_Frame, self.No_Bank_Account, "0", "1", "2", "3", "4", "5")
        self.No_Bank_Account_DropDown_Menu.grid(row = 0, column = 4, sticky = E, padx = 15, pady = 3)
        self.No_Bank_Account_DropDown_Menu.configure(font = self.fontStyle)

        self.MobileBankAccountNoText = Label(self.User_Information_Frame, text = "Number of Mobile Bank Accounts:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 0, column = 5, sticky = E)
        global No_Mobile_Bank_Account
        self.No_Mobile_Bank_Account = StringVar(self.root)
        self.No_Mobile_Bank_Account.set("")
        self.No_Mobile_Bank_Account_DropDown_Menu = OptionMenu(self.User_Information_Frame, self.No_Mobile_Bank_Account, "0", "1", "2", "3", "4", "5")
        self.No_Mobile_Bank_Account_DropDown_Menu.grid(row = 0, column = 6, sticky = E, padx = 15, pady = 3)
        self.No_Mobile_Bank_Account_DropDown_Menu.configure(font = self.fontStyle)

        self.NoPersonGivenMoneyText = Label(self.User_Information_Frame, text = "Number of Persons You've \nLent Money To:", justify = RIGHT, bg = 'white', font = self.fontStyle).grid(row = 1, column = 3, sticky = E)
        global No_Person_Given
        self.No_Person_Given = StringVar(self.root)
        self.No_Person_Given.set("")
        self.No_Person_Given_Money_Menu = OptionMenu(self.User_Information_Frame, self.No_Person_Given, "0", "1", "2", "3", "4", "5")
        self.No_Person_Given_Money_Menu.grid(row = 1, column = 4, sticky = E, padx = 15, pady = 3)
        self.No_Person_Given_Money_Menu.configure(font = self.fontStyle)

        self.NoPersonTakenMoneyText = Label(self.User_Information_Frame, text = "Number of Persons You've \nBorrowed Money From:", justify = RIGHT, bg = 'white', font = self.fontStyle).grid(row = 1, column = 5, sticky = E)
        global No_Person_Taken
        self.No_Person_Taken = StringVar(self.root)
        self.No_Person_Taken.set("")
        self.No_Person_Taken_Money_Menu = OptionMenu(self.User_Information_Frame, self.No_Person_Taken, "0", "1", "2", "3", "4", "5")
        self.No_Person_Taken_Money_Menu.grid(row = 1, column = 6, sticky = E, padx = 15, pady = 3)
        self.No_Person_Taken_Money_Menu.configure(font = self.fontStyle)


        self.EnterButton = Button(self.User_Information_Frame, text = "Enter", width = 10, background = "green", foreground = "white", command = self.Take_Banking_Data, font = self.fontStyle)
        self.EnterButton.grid(row = 0, column = 7, padx = 55, pady = 3, rowspan = 2)


    def Take_Banking_Data(self):

        self.EnterButton['state'] = 'disabled'

        #self.root = root
        # removing the user information frame from the window
        #User_Information_Frame.destroy()

        # initializing a dictionary to store all the banking information entered by the user and making it global so that it can be accessed by functions outiside this one
        global Banking_Information_Initial
        self.Banking_Information_Initial = {}

        # Child Frame 2: MobileBanking Information Frame
        self.Banking_Information_Frame = LabelFrame(self.Information_Frame, text = "Banking Information", borderwidth = 5, padx = 10, pady = 15, width = 1300, height = 170, bg = 'white', font = self.fontStyle)
        self.Banking_Information_Frame.grid(row = 1, column = 0, sticky = W)
        #Banking_Information_Frame.#grid_propagate(False)
        if int(self.No_Bank_Account.get()) != 0:
            for x in range(int(self.No_Bank_Account.get())):
                self.Banking_Information_Initial["Bank_Text_" + str(x+1) ] = Label(self.Banking_Information_Frame, text = "        Bank Account " + str(x +1) + ":", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = x + 2, column = 0, sticky = E)
                self.Banking_Information_Initial["Bank_Account_Entry_" + str(x + 1)] = Entry(self.Banking_Information_Frame, width = 34,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
                self.Banking_Information_Initial["Bank_Account_Entry_" + str(x + 1)].grid(row = x + 2, column = 1, columnspan = 2, sticky = W, padx = 10, pady = 5)
                self.Banking_Information_Initial["Bank_Balance_Text_" + str(x + 1)] = Label(self.Banking_Information_Frame, text = "Balance:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = x + 2, column = 3, sticky = W)
                self.Banking_Information_Initial["Bank_Balance_Entry_" + str(x + 1)] = Entry(self.Banking_Information_Frame, width = 34,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
                self.Banking_Information_Initial["Bank_Balance_Entry_" + str(x + 1)].grid(row = x + 2, column = 4, columnspan = 2, sticky = W, padx = 10, pady = 5)


        # Mobile Banking Information Frame(Child Frame 3)
        self.Mobile_Banking_Information_Frame = LabelFrame(self.Information_Frame, text = "Mobile Banking Information", borderwidth = 5, padx = 10, pady = 15, width = 1300, height = 170, bg = 'white', font = self.fontStyle)
        self.Mobile_Banking_Information_Frame.grid(row = 2, column = 0, sticky = W)

        if int(self.No_Mobile_Bank_Account.get()) != 0:
            for x in range(int(self.No_Mobile_Bank_Account.get())):
                self.Banking_Information_Initial["Mobile_Bank_Text_" + str(x+1) ] = Label(self.Mobile_Banking_Information_Frame, text = "Mobile Bank Account " + str(x +1) + ":", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = x + 2, column = 0, sticky = E)
                self.Banking_Information_Initial["Mobile_Bank_Account_Entry_" + str(x + 1)] = Entry(self.Mobile_Banking_Information_Frame, width = 34,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
                self.Banking_Information_Initial["Mobile_Bank_Account_Entry_" + str(x + 1)].grid(row = x + 2, column = 1, columnspan = 2, sticky = W, padx = 10, pady = 5)
                self.Banking_Information_Initial["Mobile_Bank_Balance_Text_" + str(x + 1)] = Label(self.Mobile_Banking_Information_Frame, text = "Balance:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = x + 2, column = 3, sticky = W)
                self.Banking_Information_Initial["Mobile_Bank_Balance_Entry_" + str(x + 1)] = Entry(self.Mobile_Banking_Information_Frame, width = 34,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
                self.Banking_Information_Initial["Mobile_Bank_Balance_Entry_" + str(x + 1)].grid(row = x + 2, column = 4, columnspan = 2, sticky = W, padx = 10, pady = 5)


        # Personal Loan Information Frame(Child Frame 4)
        self.Personal_Loan_Information_Frame = LabelFrame(self.Information_Frame,  text = "Personal Loan Information", labelanchor = 'n', borderwidth = 5, pady = 15, width = 130, height = 90, bg = 'white', font = self.fontStyle)
        self.Personal_Loan_Information_Frame.grid(row = 1, column = 1, sticky = NW, rowspan = 3)

        self.Given_Frame = LabelFrame(self.Personal_Loan_Information_Frame,  text = "Lent To", labelanchor = 'n', borderwidth = 5, padx = 5, pady = 15, width = 100, height = 90, bg = 'white', font = self.fontStyle)
        self.Given_Frame.grid(row = 0, column = 0, sticky = NW)

        if int(self.No_Person_Given.get()) != 0:
            position = 0
            for x in range(int(self.No_Person_Given.get())):
                self.Banking_Information_Initial["Given_Name_" + str(x+1) ] = Label(self.Given_Frame, text = "Person " + str(x +1) + ":", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = position, column = 0, sticky = E)
                self.Banking_Information_Initial["Given_Name_Entry_" + str(x + 1)] = Entry(self.Given_Frame, width = 34,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
                self.Banking_Information_Initial["Given_Name_Entry_" + str(x + 1)].grid(row = position, column = 1, columnspan = 2, sticky = W, padx = 10, pady = 5)
                self.Banking_Information_Initial["Given_Amount_Text_" + str(x + 1)] = Label(self.Given_Frame, text = "Amount:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = position + 1, column = 0, sticky = W)
                self.Banking_Information_Initial["Given_Amount_Entry_" + str(x + 1)] = Entry(self.Given_Frame, width = 34,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
                self.Banking_Information_Initial["Given_Amount_Entry_" + str(x + 1)].grid(row = position + 1, column = 1, columnspan = 2, sticky = W, padx = 10, pady = 5)
                position += 2


        self.Taken_Frame = LabelFrame(self.Personal_Loan_Information_Frame,  text = "Borrowed From", labelanchor = 'n', borderwidth = 5, pady = 15, width = 100, height = 90, bg = 'white', font = self.fontStyle)
        self.Taken_Frame.grid(row = 1, column = 0, sticky = NW)

        if int(self.No_Person_Taken.get()) != 0:
            pos = 0
            for x in range(int(self.No_Person_Taken.get())):
                self.Banking_Information_Initial["Taken_Name_" + str(x+1) ] = Label(self.Taken_Frame, text = "Person " + str(x +1) + ":", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = position, column = 0, sticky = E)
                self.Banking_Information_Initial["Taken_Name_Entry_" + str(x + 1)] = Entry(self.Taken_Frame, width = 34,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
                self.Banking_Information_Initial["Taken_Name_Entry_" + str(x + 1)].grid(row = position, column = 1, columnspan = 2, sticky = W, padx = 10, pady = 5)
                self.Banking_Information_Initial["Taken_Amount_Text_" + str(x + 1)] = Label(self.Taken_Frame, text = "Amount:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = position + 1, column = 0, sticky = W)
                self.Banking_Information_Initial["Taken_Amount_Entry_" + str(x + 1)] = Entry(self.Taken_Frame, width = 34,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
                self.Banking_Information_Initial["Taken_Amount_Entry_" + str(x + 1)].grid(row = position + 1, column = 1, columnspan = 2, sticky = W, padx = 10, pady = 5)
                position += 2


        # Child Frame 5 :  Cash Information Frame
        self.CashInformation_Frame = LabelFrame(self.Information_Frame, text = "Cash Information", borderwidth = 5, padx = 10, pady = 15, width = 1300, height = 170, bg = 'white', font = self.fontStyle)
        self.CashInformation_Frame.grid(row = 3, column = 0, sticky = W, ipadx = 202)
        #self.CashInformation_Frame.#grid_propagate(False)

        self.CashText = Label(self.CashInformation_Frame, text = "          Cash Balance:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 0, column = 0, sticky = W)
        global CashEntry
        self.CashEntry = Entry(self.CashInformation_Frame, width = 34,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
        self.CashEntry.insert(0, "0")
        self.CashEntry.grid(row = 0, column = 1, columnspan = 2, sticky = W, padx = 10, pady = 5)


        # Submit button will start the entry frame and allow user to input the data and see statistics about the entered data
        self.SubmitButton = Button(self.Information_Frame, text = "Submit", width = 154, background = "green", foreground = "white", command = self.Open_Balance_Sheet, font = self.fontStyle).grid(row = 5, column = 0, padx = 3, pady = 10, columnspan = 2, sticky = N)


        return

    def Open_Balance_Sheet(self):

        #Storing the user input bank account names into a list to help naming the columns of the Balance Sheet
        #global Bank_Account_Entry_1, Bank_Account_Entry_2, Bank_Account_Entry_3, MobileBank_Account_Entry_1, MobileBank_Account_Entry_2, MobileBank_Account_Entry_3
        #global Bank_Balance_Entry1, Bank_Balance_Entry2, Bank_Balance_Entry3, MobileBank_Balance_Entry1, MobileBank_Balance_Entry2, MobileBank_Balance_Entry3,
        global CashEntry
        global Banking_Information_Initial
        self.Bank_Column_Names = []
        self.Data_Sheet_First_Row_Banks = []
        if int(self.No_Bank_Account.get()) != 0:
            for x in range(int(self.No_Bank_Account.get())):
                self.Bank_Column_Names.append(self.Banking_Information_Initial["Bank_Account_Entry_" + str(x + 1)].get())
                self.Data_Sheet_First_Row_Banks.append( self.Banking_Information_Initial["Bank_Balance_Entry_" + str(x + 1)].get())
        if int(self.No_Mobile_Bank_Account.get()) != 0:
            for x in range(int(self.No_Mobile_Bank_Account.get())):
                self.Bank_Column_Names.append(self.Banking_Information_Initial["Mobile_Bank_Account_Entry_" + str(x + 1)].get())
                self.Data_Sheet_First_Row_Banks.append( self.Banking_Information_Initial["Mobile_Bank_Balance_Entry_" + str(x + 1)].get())

        #Bank_Column_Names = [Banking_Information_Initial["Bank_Account_Entry_1"].get(), Banking_Information_Initial["Bank_Account_Entry_2"].get(), Banking_Information_Initial["Bank_Account_Entry_3"].get(), Banking_Information_Initial["Mobile_Bank_Account_Entry_1"].get(), Banking_Information_Initial["Mobile_Bank_Account_Entry_2"].get(), Banking_Information_Initial["Mobile_Bank_Account_Entry_3"].get()]#MobileBank_Account_Entry_1.get(), MobileBank_Account_Entry_2.get(),MobileBank_Account_Entry_3.get()]
        #Data_Sheet_First_Row_Banks = [Banking_Information_Initial["Bank_Balance_Entry_1"].get(), Banking_Information_Initial["Bank_Balance_Entry_2"].get(), Banking_Information_Initial["Bank_Balance_Entry_3"].get(), Banking_Information_Initial["Mobile_Bank_Balance_Entry_1"].get(), Banking_Information_Initial["Mobile_Bank_Balance_Entry_2"].get(), Banking_Information_Initial["Mobile_Bank_Balance_Entry_3"].get()]#MobileBank_Balance_Entry1.get(), MobileBank_Balance_Entry2.get(), MobileBank_Balance_Entry3.get()]

        # Removing the empty inputs from the lists
        self.Bank_Column_Names = [string for string in self.Bank_Column_Names if string != ""]
        self.Data_Sheet_First_Row_Banks = [string for string in self.Data_Sheet_First_Row_Banks if string != ""]
        self.Data_Sheet_First_Row_Banks = [int(i) for i in self.Data_Sheet_First_Row_Banks]

        today = str(_datetime.date.today())
        today = datetime.strptime(today, "%Y-%m-%d")
        today = today.strftime("%d-%b-%Y")

        # Column names for columns that will have the transaction details
        self.Transaction_Column_Names = ["Date", "Direction", "Tag", "Item", "Amount", "Price", "Total System In", "Total System Out", "Cash Balance"]
        self.DataSheet_First_Row_Transaction = [str(today),"","","","","",0,0, int(self.CashEntry.get())]


        # Column Names for Given and Taken
        self.Given_Taken_Column_Names = ["Given", "Taken"]
        self.Data_Sheet_First_Row_Given_Taken = [ 0, 0 ]

        # Combining all three lists for column names
        self.Sheet_Column_Names = self.Transaction_Column_Names + self.Bank_Column_Names + self.Given_Taken_Column_Names
        self.Data_Sheet_Rows = [ self.DataSheet_First_Row_Transaction + self.Data_Sheet_First_Row_Banks + self.Data_Sheet_First_Row_Given_Taken ]


        Total_Given = 0
        if int(self.No_Person_Given.get()) != 0:
            for x in range(int(self.No_Person_Given.get())):
                Total_Given = Total_Given + int(self.Banking_Information_Initial["Given_Amount_Entry_" + str(x + 1)].get())
                #self.Data_Sheet_Rows =
                self.Data_Sheet_Rows.append([str(today),"Given-Cash Balance","Money", self.Banking_Information_Initial["Given_Name_Entry_" + str(x + 1)].get(),"",int(self.Banking_Information_Initial["Given_Amount_Entry_" + str(x + 1)].get()),0,0, int(self.CashEntry.get())] + self.Data_Sheet_First_Row_Banks + [ Total_Given, 0])
        global Total_Taken
        Total_Taken = 0
        if int(self.No_Person_Taken.get()) != 0:
            for x in range(int(self.No_Person_Taken.get())):
                Total_Taken = Total_Taken + int(self.Banking_Information_Initial["Taken_Amount_Entry_" + str(x + 1)].get())
                #self.Data_Sheet_Rows =
                self.Data_Sheet_Rows.append([str(today),"Taken-Cash Balance","Money", self.Banking_Information_Initial["Taken_Name_Entry_" + str(x + 1)].get(),"",int(self.Banking_Information_Initial["Taken_Amount_Entry_" + str(x + 1)].get()),0,0, int(self.CashEntry.get())] + self.Data_Sheet_First_Row_Banks + [ Total_Given, Total_Taken])


            #self.Bank_Column_Names.append(self.Banking_Information_Initial["Given_Amount_Entry_" + str(x + 1)].get())
            #self.Data_Sheet_First_Row_Banks.append( self.Banking_Information_Initial["Bank_Balance_Entry_" + str(x + 1)].get())

        # Creating a pnadas dataframe with first row all equal to zeros for initialization of data input system
        self.Initializing_Data_Sheet = pd.DataFrame(data = self.Data_Sheet_Rows, columns = self.Sheet_Column_Names)

        # Saving the pandas dataframe to Microsoft Excel file for initializing the transaction input the first time for the user
        # creating excel writer
        global writer
        self.writer = pd.ExcelWriter('Balance Sheet.xlsx', mode = 'w', engine = "openpyxl")
        self.Initializing_Data_Sheet.to_excel(self.writer, sheet_name = "Sheet", index = None)

        '''
        # newly created frames for transaction entry and statistics will be over
        # this tags frame if not destroyed. So, deleting those widgets for making
        # sure that those tag entry widgets are not visible from back
        for widgets in self.root.winfo_children():
            widgets.destroy()
        '''

        from Taking_Tag_Choice import Tag_Options_User_Choice
        Tag_Options_User_Choice_Window = Tag_Options_User_Choice(root = self.root, Name_Entry = self.Name_Entry, writer = self.writer)
        Tag_Options_User_Choice_Window.start()


    def start(self):
        self.root.mainloop()
