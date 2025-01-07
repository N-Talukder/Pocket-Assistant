# importing required packages
from tkinter import *
import tkinter.font as tkFont
from datetime import datetime
import calendar
from tkcalendar import *
import _datetime
import openpyxl
import pathlib
import pandas as pd
import babel.numbers

class Transaction_Entry:
    def __init__(self, root):

        self.TransactionFrame = root
        self.root = self.TransactionFrame

        self.fontStyle = tkFont.Font(family="Comic Sans MS", size=11, weight = "bold")

        self.TransactionButton = Button(self.TransactionFrame, text = "Enter New Transaction", padx = 203, pady = 10, command = self.Transaction_Details, font = self.fontStyle)
        self.TransactionButton.grid(row = 0, column = 0, padx = 10, pady = 10, sticky = W + E, columnspan = 7)
        self.TransactionButton.grid_propagate(False)


        # when the user cllicks on Enter a new transaction, we will use this function to display options regarding the traansaction
    def Transaction_Details(self):
        # creating label widgets for user understanding what to enter in the box next to it and storing those inside a variable
        #Date Input Creation
        DateText = Label(self.TransactionFrame, text = "Date: ", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 1, column = 0, sticky = E)
        #designing a frame for calendar and putting it up in the transactionframe window
        CalendarFrame = LabelFrame(self.TransactionFrame, borderwidth = 5, padx = 5, pady = 5, width = 400, height = 210, bg = 'white')
        CalendarFrame.grid(row = 1, column = 1 , padx = 10, pady = 3, columnspan = 5, sticky = W)
        CalendarFrame.grid_propagate(False)
        #getting current date from interacting with the computer for setting it as the default when the calendar is called for taking the user input
        today = str(_datetime.date.today())
        today = datetime.strptime(today, "%Y-%m-%d")
        #putting the calendar up inside the created frame calendarframe
        global cal
        cal = Calendar(CalendarFrame, selectmode = "day", year = today.year, day = today.day, month = today.month, width = 350)
        cal.grid(row = 0, column = 0, sticky = W, ipadx = 50, ipady = 0, padx = 14)
        cal.grid_propagate(False)

        #Getting the Item name from the user
        ItemText = Label(self.TransactionFrame, text = "Item/Person Name:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 2, column = 0, sticky = E)
        global I
        I = Entry(self.TransactionFrame, width = 43,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
        I.grid(row = 2, column = 1, columnspan = 6, sticky = W, padx = 10, pady = 3)

        #Getting the Amount from user
        AmountText = Label(self.TransactionFrame, text = "Amount/Quantity:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 3, column = 0, sticky = E)
        global A
        A = Entry(self.TransactionFrame, width = 30, borderwidth = 3, foreground = "black", background = "white", font = self.fontStyle)
        A.grid(row = 3, column = 1, sticky = W, padx = 10, pady = 3, columnspan = 1)
        global Unit
        Unit = StringVar(self.root)
        Unit.set("")
        UnitDropDownMenu = OptionMenu(self.TransactionFrame, Unit, "pcs", "kg", "g", "L", "mL", "bottle", "pack", "month", "")
        UnitDropDownMenu.grid(row = 3, column = 2, columnspan = 2, sticky = E, padx = 15, pady = 3)#column = 2, sticky = W, padx = 13)
        UnitDropDownMenu.configure(font = self.fontStyle)

        # Information about the money involved either spent or system in
        PriceText = Label(self.TransactionFrame, text = "Price/Cost:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 4, column = 0, sticky = E)
        global P
        P = Entry(self.TransactionFrame, width = 43, borderwidth = 3, foreground = "black", background = "white", font = self.fontStyle)
        P.grid(row = 4, column = 1, columnspan = 6, sticky = W, padx = 10, pady = 3)

        #Tags that will be used to divide the money involved into separate sectors when requested by the user through the options in the StatisticsFrame
        TagText = Label(self.TransactionFrame, text = "Transaction Tag:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 5, column = 0, sticky = E)
        global Tag, Tags
        Tag = StringVar(self.root)
        Tag.set("")
        #reading user input tags from excel file created when user input the tags preferences the first time
        Tags = pd.read_excel("Balance Sheet.xlsx", sheet_name = "Tags", skiprows = 0, header = None)#squeeze = True,
        Tags = [Tag[0] for Tag in Tags.values]#Tags.values[0].tolist()

        TagDropdownMenu = OptionMenu(self.TransactionFrame, Tag, *Tags)
        TagDropdownMenu.grid(row = 5, column = 1, columnspan = 2, sticky = W, padx = 10, ipadx = 82)
        TagDropdownMenu.configure(font = self.fontStyle)

        #Directions help classify the informations input about the transaction and also helps in the subsequent calculations for ledger management
        DirectionText = Label(self.TransactionFrame, text = "Direction of Transaction:", bg = 'white', font = self.fontStyle).grid(row = 6, column = 0, sticky = E)

        #reading user bank accounts from excel file created when user input the bank details the first time
        Accounts = pd.read_excel("Balance Sheet.xlsx", sheet_name = "Sheet", header = None).iloc[0]#squeeze = True, 
        Accounts = Accounts.values.tolist()
        del Accounts[ 0 : 8 ]
        del Accounts[ (len(Accounts)-2) : len(Accounts) ]

        InFrame = LabelFrame(self.TransactionFrame, borderwidth = 5, padx = 5, width = 400, height = 50, bg = 'white')
        InFrame.grid(row = 6, column = 1 , padx = 10, pady = 5, columnspan = 3, sticky = E)
        InFrame.grid_propagate(False)

        def DirectionFinal( button_number, direction):
            DirectionButtonDictionary["DirectionCheck" + str(button_number)].configure(bg = 'green')
            global D
            D = direction

            for x in range(1, 6):
                if x != button_number:
                    DirectionButtonDictionary["DirectionCheck" + str(x)].configure(bg = 'white')
            return
        global DirectionButtonDictionary
        DirectionButtonDictionary = {}

        DirectionButtonDictionary['DirectionCheck1'] = Button(InFrame, text = 'In', width = 6, padx = 10, background = "white", foreground = "black", font = self.fontStyle, command = lambda: DirectionFinal(1, "In"))
        DirectionButtonDictionary['DirectionCheck1'].grid(row = 0, column = 0, sticky = W, padx = 10)
        DirectionButtonDictionary['DirectionCheck2'] = Button(InFrame, text = 'Taken', width = 6, padx = 10, background = "white", foreground = "black", font = self.fontStyle, command = lambda: DirectionFinal(2, "Taken"))
        DirectionButtonDictionary['DirectionCheck2'].grid(row = 0, column = 1, sticky = W, padx = 10)
        global InBank
        InBank = StringVar(self.root)
        InBank.set("")
        InMenu = OptionMenu(InFrame, InBank, *Accounts)
        InMenu.grid(row = 0, column = 4, sticky = E, padx = 10, pady = 0, columnspan = 2, ipadx = 8)
        InMenu.configure(font = self.fontStyle, bg ='white')

        OutFrame = LabelFrame(self.TransactionFrame, borderwidth = 5, padx = 5, width = 400, height = 50, bg = 'white')
        OutFrame.grid(row = 7, column = 1 , padx = 10, pady = 5, columnspan = 3, sticky = E)
        OutFrame.grid_propagate(False)

        DirectionButtonDictionary['DirectionCheck3'] = Button(OutFrame, text = 'Out', width = 6, padx = 10, background = "white", foreground = "black", font = self.fontStyle, command = lambda: DirectionFinal(3, "Out"))
        DirectionButtonDictionary['DirectionCheck3'].grid(row = 0, column = 0, sticky = W, padx = 10)
        DirectionButtonDictionary['DirectionCheck4'] = Button(OutFrame, text = 'Given', width = 6,  padx = 10, background = "white", foreground = "black", font = self.fontStyle, command = lambda: DirectionFinal(4, "Given"))
        DirectionButtonDictionary['DirectionCheck4'].grid(row = 0, column = 1, sticky = W, padx = 10)
        global OutBank
        OutBank = StringVar(self.root)
        OutBank.set("")
        OutMenu = OptionMenu(OutFrame, OutBank, *Accounts)
        OutMenu.grid(row = 0, column = 4, sticky = E, padx = 10, pady = 0, columnspan = 2, ipadx = 8)
        OutMenu.configure(font = self.fontStyle, bg ='white')

        ConversionFrame = LabelFrame(self.TransactionFrame, borderwidth = 5, padx = 5, width = 400, height = 50, bg = 'white')
        ConversionFrame.grid(row = 8, column = 1 , padx = 10, columnspan = 3, sticky = E, pady = 7)
        ConversionFrame.grid_propagate(False)

        DirectionButtonDictionary['DirectionCheck5'] = Button(ConversionFrame, text = 'Conversion', width = 6,  padx = 10, background = "white", foreground = "black", font = self.fontStyle, command = lambda: DirectionFinal(5, "Conversion"))
        DirectionButtonDictionary['DirectionCheck5'].grid(row = 0, column = 1, sticky = W, padx = 10)
        global FromBank
        FromBank = StringVar(self.root)
        FromBank.set("From")
        FromMenu = OptionMenu(ConversionFrame, FromBank, *Accounts)
        FromMenu.grid(row = 0, column = 2, sticky = E, padx = 10, pady = 0, columnspan = 2)
        FromMenu.configure(font = self.fontStyle, bg ='white')
        global ToBank
        ToBank = StringVar(self.root)
        ToBank.set("To")
        ToMenu = OptionMenu(ConversionFrame, ToBank, *Accounts)
        ToMenu.grid(row = 0, column = 4, sticky = E, padx = 10, pady = 0, columnspan = 2)
        ToMenu.configure(font = self.fontStyle, bg ='white')

        # creating a submit button for taking input of the transaction with the user's permission
        SubmitButton = Button(self.TransactionFrame, text = "Submit", padx = 180, pady = 5, background = "green", foreground = "white", command = self.Data_Check, font = self.fontStyle).grid(row = 11, column = 0, padx = 10, sticky = W + E, columnspan = 5)

    # this function is for checking if the transaction entered is after the last entry in the excel file or it's in a back date
    # this will also
    def Data_Check(self):

        global Item, Amount, Price, T, Direction, D
        Item = I.get()
        Amount = A.get() + " " + Unit.get()
        Price = P.get()
        T = Tag.get()

        global Date
        Date = cal.get_date()
        Date = datetime.strptime(Date, "%m/%d/%y")

        if D == "In" or D == "Taken":
            global InBank
            y = InBank.get()

            Direction = D + "-" + y

        elif D == "Out" or D == "Given":
            global OutBank
            y = OutBank.get()

            Direction = D + "-" + y

        elif D == "Conversion":
            global FromBank, ToBank
            y1 = FromBank.get()
            y2 = ToBank.get()

            Direction = D + ": " + y1 + "-> " + y2

        if Price == "" or T == "":
            return

        else:
            #clearing the dropdownmenu selections and text boxes after storing the results into excel
            I.delete(0,"end")
            A.delete(0,"end")
            P.delete(0,"end")
            Tag.set("")
            Unit.set("")

            global DirectionButtonDictionary
            for x in range(1, 6):
                DirectionButtonDictionary["DirectionCheck" + str(x)].configure(bg = 'white')

            #reading the already existing excel data sheet into a python data frame for calculations
            filePath = pathlib.Path('Balance Sheet.xlsx')
            global dataFrameEntry
            dataFrameEntry = pd.read_excel(filePath, sheet_name='Sheet', header = 0)
            dataFrameEntry['Date'] = pd.to_datetime(dataFrameEntry['Date'])

            def time_format_change(a):
                a = datetime.strptime(str(a), "%Y-%m-%d %H:%M:%S")
                a = datetime.strftime(a, "%d-%b-%Y")
                return a

            if Date >= dataFrameEntry['Date'].iloc[len(dataFrameEntry)-1]:


                dataFrameEntry['Date'] = dataFrameEntry['Date'].apply(lambda x: time_format_change(x))

                New_Row_Transaction_Information = [Date.strftime("%d-%b-%Y"), Direction, T[2:-3], Item, Amount, float(Price)]#int(Price)

                ExcelNewLine = self.Excel_Calculation(dataFrameEntry, New_Row_Transaction_Information)

                #dataFrameEntry = dataFrameEntry.append(ExcelNewLine, ignore_index = True)
                dataFrameEntry = pd.concat([dataFrameEntry, pd.DataFrame([ExcelNewLine])], ignore_index=True)

                self.Exporting_Transaction_To_Excel(dataFrameEntry)

            else:

                dataFrameEntry_Left_After_Entered_Transactiom = dataFrameEntry[ dataFrameEntry['Date'] > Date ][['Date', 'Direction', 'Tag', 'Item', 'Amount', 'Price']]

                dataFrameEntry = dataFrameEntry[ dataFrameEntry['Date'] <= Date ]

                dataFrameEntry['Date'] = dataFrameEntry['Date'].apply(lambda x: time_format_change(x))

                New_Row_Transaction_Information = [Date.strftime("%d-%b-%Y"), Direction, T, Item, Amount, Price]


                ExcelNewLine = self.Excel_Calculation(dataFrameEntry, New_Row_Transaction_Information)

                #dataFrameEntry = dataFrameEntry.append(ExcelNewLine, ignore_index = True)
                dataFrameEntry = pd.concat([dataFrameEntry, pd.DataFrame([ExcelNewLine])], ignore_index=True)


                for x in range(len(dataFrameEntry_Left_After_Entered_Transactiom)):

                    Date = dataFrameEntry_Left_After_Entered_Transactiom['Date'].iloc[x]
                    Direction = dataFrameEntry_Left_After_Entered_Transactiom['Direction'].iloc[x]
                    T = dataFrameEntry_Left_After_Entered_Transactiom['Tag'].iloc[x]
                    #print(T)
                    Item = dataFrameEntry_Left_After_Entered_Transactiom['Item'].iloc[x]
                    Amount = dataFrameEntry_Left_After_Entered_Transactiom['Amount'].iloc[x]
                    Price = dataFrameEntry_Left_After_Entered_Transactiom['Price'].iloc[x]


                    New_Row_Transaction_Information = [Date.strftime("%d-%b-%Y"), Direction, T, Item, Amount, float(Price)]#int(Price)

                    ExcelNewLine = self.Excel_Calculation(dataFrameEntry, New_Row_Transaction_Information)

                    #dataFrameEntry = dataFrameEntry.append(ExcelNewLine, ignore_index = True)
                    dataFrameEntry = pd.concat([dataFrameEntry, pd.DataFrame([ExcelNewLine])], ignore_index=True)

                self.Exporting_Transaction_To_Excel(dataFrameEntry)

        return



    # this function does the calculations for the user input transaction using the options selected by the user about the transaction
    def Excel_Calculation(self, dataFrameEntry, New_Row_Transaction_Information):

        #Date = New_Row_Transaction_Information[1]
        Price = float(New_Row_Transaction_Information[5])

        ExcelNewLine = {}

        Excel_Sheet_Column_Names = pd.read_excel("Balance Sheet.xlsx", sheet_name = "Sheet", header = None).iloc[0]#squeeze = True, 
        Excel_Sheet_Column_Names = Excel_Sheet_Column_Names.values.tolist()

        for x in Excel_Sheet_Column_Names[ 6 : ( len(Excel_Sheet_Column_Names)) ]:

            ExcelNewLine[ x ] = float(dataFrameEntry[x][len(dataFrameEntry)-1])

        if "In" in Direction:
            y = Direction.replace("In-", "")
            ExcelNewLine[ y ] = ExcelNewLine[ y ] + Price
            ExcelNewLine[ "Total System In" ] = ExcelNewLine["Total System In"] + Price
        elif "Taken" in Direction:
            y = Direction.replace("Taken-", "")
            ExcelNewLine[ "Taken" ] = ExcelNewLine["Taken"] + Price
            ExcelNewLine[ y ] = ExcelNewLine[ y ] + Price
        elif "Out" in Direction:
            y = Direction.replace("Out-", "")
            ExcelNewLine[ y ] = ExcelNewLine[ y ] - Price
            ExcelNewLine[ "Total System Out" ] = ExcelNewLine["Total System Out"] + Price
        elif "Given" in Direction:
            y = Direction.replace("Given-", "")
            ExcelNewLine[ y ] = ExcelNewLine[ y ] - Price
            ExcelNewLine[ "Given" ] = ExcelNewLine["Given"] + Price
        elif "Conversion" in Direction:
            y = Direction.split("-> ")
            y1 = y[0].replace("Conversion: ", "")
            y2 = y[1]
            ExcelNewLine[ y1 ] = ExcelNewLine[ y1 ] - Price
            ExcelNewLine[ y2 ] = ExcelNewLine[ y2 ] + Price

        r = 0
        for x in Excel_Sheet_Column_Names[ 0 : 6 ]:
            ExcelNewLine[ x ] = New_Row_Transaction_Information[r]
            r += 1

        return ExcelNewLine

    # funciton for exporting the transaction details entered to excel file
    def Exporting_Transaction_To_Excel(self, dataFrameEntry):

        # creating excel writer for storing the calculation outputs into different excel sheets
        writer = pd.ExcelWriter('Balance Sheet.xlsx', mode = 'a', if_sheet_exists = 'replace', engine = "openpyxl")

        # writing the transactions within user input dates into a new excel sheet
        dataFrameEntry.to_excel(writer, sheet_name = "Sheet", index = None)

        writer.close()

        return


    def start(self):
        self.root.mainloop()
