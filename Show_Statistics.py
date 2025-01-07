from tkinter import *
import tkinter.font as tkFont
from datetime import datetime
import calendar
from tkcalendar import *
import _datetime
import openpyxl
import pathlib
import pandas as pd
from openpyxl.chart import BarChart, Reference
from tkinter import messagebox
import babel.numbers

class Show_Statistics:
    def __init__(self, root):

        self.StatisticsFrame = root

        self.fontStyle = tkFont.Font(family="Comic Sans MS", size=11, weight = "bold")

        self.StatisticsButton = Button(self.StatisticsFrame, text = "See Statistics", width = 13, padx = 292, pady = 10, command = self.Statistics_Options, font = self.fontStyle)
        self.StatisticsButton.grid(row = 0, column = 0, padx = 10, sticky = E, pady = 8, columnspan = 9)
        self.StatisticsButton.grid_propagate(False)

    # this function calculates statistical values for displaying them in the respective sheets in the excel file
    def Statistics_Calculations(self):

        #collecting transaction details submitted into variables
        global startDate, endDate, Item, T, Direction, D, Account

        startDate = self.startCal.get_date()
        endDate = self.endCal.get_date()
        Item = self.I.get()
        T = self.Tag.get()
        #Direction = D

        # reading the entire excel datasheet with all user input transaction details into pandas dataframe as this form has better functions for complex relevant tasks
        filePath = pathlib.Path('Balance Sheet.xlsx')
        dataFrame = pd.read_excel(filePath, sheet_name='Sheet', header = 0)

        # converting the dates in the first column in python datetime objects for later operations
        dataFrame['Date'] = pd.to_datetime(dataFrame['Date'], format = "%d-%b-%Y")

        # Selecting the rows between two dates input by the user for showing statistics withing this date limits
        DataInDateRange = dataFrame[dataFrame['Date'].isin(pd.date_range(startDate, endDate))]

        # creating excel writer for storing the calculation outputs into different excel sheets
        writer = pd.ExcelWriter('Balance Sheet.xlsx', mode = 'a', if_sheet_exists = 'replace', engine = "openpyxl")

        # for better representation in the excel file, changing the datetime object type to string type
        DataInDateRange['Date'] = DataInDateRange['Date'].astype(str)
        # writing the transactions within user input dates into a new excel sheet
        DataInDateRange.to_excel(writer, sheet_name = "Data within Dates", index = None)

        #calculating the total system in within the selected dates
        total_System_In = DataInDateRange[DataInDateRange['Direction'].str.contains("In", na = False)]['Price'].sum(axis = None, skipna = True)

        global Tags
        total_System_Out = 0
        tag_order_wise_totals_within_dates = []
        for x in self.Tags:
            x = str(x)
            a = DataInDateRange[DataInDateRange['Tag'].str.contains(x, na = False)]['Price'].sum(axis = None, skipna = True)
            tag_order_wise_totals_within_dates.append(a)
            total_System_Out = total_System_Out + a


        Total__In_Out = pd.DataFrame([["Total In", total_System_In], ["", ""], ["Total Out", total_System_Out], ["", ""]], columns=None, index=None)

        #writerappend = pd.ExcelWriter('Balance Sheet.xlsx', mode = 'a', if_sheet_exists = 'replace', engine = "openpyxl")
        tag_order_wise_totals_within_dates = pd.DataFrame(tag_order_wise_totals_within_dates, index = None)

        Tags = pd.read_excel("Balance Sheet.xlsx", sheet_name = "Tags", skiprows = 0, header = None)#squeeze = True, 
        Tags = Tags.values.tolist()
        Tags = pd.DataFrame(Tags, index = None)
        Tag_plus_totals = pd.concat([ Tags, tag_order_wise_totals_within_dates], axis=1, join='inner', ignore_index = True, )
        #Total__In_Out = Total__In_Out.append(Tag_plus_totals)
        Total__In_Out = pd.concat([Total__In_Out, Tag_plus_totals], ignore_index=True)
        Total__In_Out.to_excel(writer, sheet_name = "Cost Breakdown", index = None, startrow = 0, startcol = 0, header = None)


        # the following conditionals are for calculating totals in every sector for Item/Tag/Direction
        if Item != "":

            # if the user wants to see the data related to something he/she spent money on
            if self.personVar.get() == 0:
                # Select the rows containing the rows with the specific search string inserted for Item name
                ItemData = DataInDateRange[DataInDateRange['Item'].str.contains(Item, na = False, case = False)]
                # truncating the columns for only relevant data displaying
                ItemData = ItemData.iloc[:, [0,1,2,3,4,5]]
                #adding a new empty row to the dataframe
                ItemData.loc[ItemData.shape[0]+1] = [None, None, None, None, None, None]
                # calculating the total and storing it in the last row of the dataframe
                ItemData.iloc[[len(ItemData)-1], 4] = "Total"
                ItemData.iloc[[len(ItemData)-1], 5] = ItemData['Price'].sum(axis = None, skipna = True)
                x = "Item Data"

            # displyaing options will be different for persons compared to items as given and taken are relevant for persons.
            elif self.personVar.get() == 1:

                # Select the rows containing the rows with the person name or part of the name
                ItemData = DataInDateRange[DataInDateRange['Item'].str.contains(Item, na = False, case = False)]
                ItemData = ItemData.iloc[:, [0,1,2,3,5]]
                # calculating the totals of given and taken amount for the person of interest
                GivenBalance = ItemData[ItemData['Direction'].str.contains('Given', na = False)]['Price'].sum(axis = None, skipna = True)
                TakenBalance = ItemData[ItemData['Direction'].str.contains('Taken', na = False)]['Price'].sum(axis = None, skipna = True)
                Balance = GivenBalance - TakenBalance
                ItemData.loc[ItemData.shape[0]+1] = [None, "", None, None, None]
                #
                if Balance > 0:
                    dialogue = "She/He owes you = "
                elif Balance < 0:
                    dialogue = "You owe him/her = "
                    Balance = 0 - Balance
                else:
                    dialogue = "Total"
                print(ItemData)
                # inserting the calculated results and respective text into the dataframe last row
                ItemData.iloc[[len(ItemData)-1], 3] = dialogue
                ItemData.iloc[[len(ItemData)-1], 4] = Balance

                print(ItemData)

                x = "Person Data"
            # for better representation in the excel file, changing the datetime object type to string type
            ItemData['Date'] = ItemData['Date'].astype(str)
            # displaying the results in the excel sheet for item/person
            ItemData.to_excel(writer, sheet_name = x, index = None)

        elif T != "":

            # Select the rows containing the rows with the specific tag, calculating total and displaying those in the last row in dataframe
            TagData = DataInDateRange[DataInDateRange['Tag'].str.contains(T, na = False)]
            TagData = TagData.iloc[:, [0,1,2,3,4,5]]
            TagData.loc[TagData.shape[0]+1] = [None, None, None, None, None, None]
            TagData.iloc[[len(TagData)-1], 4] = "Total"
            TagData.iloc[[len(TagData)-1], 5] = TagData['Price'].sum(axis = None, skipna = True)

            # for better representation in the excel file, changing the datetime object type to string type
            TagData['Date'] = TagData['Date'].astype(str)
            # dataframe output to excel
            TagData.to_excel(writer, sheet_name = "Tag Data", index = None)

        if str(self.Account.get()) != "":

            # Select the rows containing the rows with the specific direction, calculating total and displaying those in the last row in the dataframe
            AccountData = DataInDateRange[DataInDateRange['Direction'].str.contains(str(self.Account.get()), na = False)]
            AccountData = AccountData[["Date", "Direction", "Tag", "Item", "Price", str(self.Account.get())]]

            AccountData["In"] = int(0)
            AccountData["Out"] = int(0)

            AccountDataInColumn = []#AccountData["In"]
            AccountDataOutColumn = []#AccountData["Out"]

            def changing_in_out_column_values(direction, price):
                if "In" in direction or ("> " + str(self.Account.get())) in direction or "Taken" in direction:
                    AccountDataInColumn.append(price)
                    AccountDataOutColumn.append(0)
                elif "Out" in direction or (str(self.Account.get() + "-> ")) in direction or "Given" in direction:
                    AccountDataOutColumn.append(price)
                    AccountDataInColumn.append(0)
                return

            AccountData.apply(lambda row: changing_in_out_column_values(row['Direction'], row['Price']), axis=1)

            AccountData["In"] = AccountDataInColumn
            AccountData["Out"] = AccountDataOutColumn


            AccountData = AccountData[["Date", "Direction", "Tag", "Item", "Price", "In", "Out", str(self.Account.get())]]
            AccountData.columns = ["Date", "Direction", "Tag", "Item", "Price", "In", "Out", str(self.Account.get()) + " Balance"]

            # for better representation in the excel file, changing the datetime object type to string type
            AccountData['Date'] = AccountData['Date'].astype(str)
            #dataframe output to excel
            AccountData.to_excel(writer, sheet_name = "Account Data", index = None)
        if str(self.Direction.get()) != "":
            print("nothing" + str(self.Account.get()))
            if str(self.Account.get()) != "":
                DirectionData = AccountData[AccountData['Direction'].str.contains(str(self.Direction.get()), na = False)]
                #AccountData = AccountData[["Date", "Direction", "Tag", "Item", "Amount", "Price", str(self.self.Account.get())]]
                title = "(Account+Direction) Data"
            else:
                DirectionData = DataInDateRange[DataInDateRange['Direction'].str.contains(str(self.Direction.get()), na = False)]

                title = "Direction Data"

            # for better representation in the excel file, changing the datetime object type to string type
            DirectionData['Date'] = DirectionData['Date'].astype(str)
            #dataframe output to excel
            DirectionData.to_excel(writer, sheet_name = title, index = None)

        # saving the excel file after the changes made
        writer.close()#save()


        # generating a graph for tag-wise comparison using bar chart created in the Cost breakdown excel sheet
        # opening a handle for taking out data and inserting a graph based on that data
        file = openpyxl.load_workbook('Balance Sheet.xlsx')
        sheet = file["Cost Breakdown"]
        #initializing and giving the data input to the chart
        chart = BarChart()
        data = Reference(sheet, min_col=2, min_row = 5, max_row = 4 + len(Tags), max_col=2)
        category = Reference(sheet, min_col=1, min_row = 6, max_row = 4 + len(Tags), max_col=1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(category)
        #stylizing the chart
        chart.style = 10
        chart.title = "Sectors and Expenditures"
        chart.y_axis.title = 'Spent Money(BDT)'
        chart.x_axis.title = 'Sectors'
        chart.shape = 4
        chart.height = 12 # default is 7.5
        chart.width = 30 # default is 15
        chart.legend = None
        # adding the chart to the sheet
        sheet.add_chart(chart, 'F3')
        #saving the file so that the added chart is not lost
        file.save('Balance Sheet.xlsx')
        # curating message for the user about where to find the results
        message = "All the transactions within the selected date interval are added to a new sheet named 'Data within Dates' in 'Balance Sheet.xlsx' file"
        if Item != "":
            if self.personVar.get() == 0:
                message += "\nA new excel sheet titled 'Item Data' is addded to the 'Balance Sheet.xlsx' file with the Item transactions requested within selected dates"

            elif self.personVar.get() == 1:
                message += "\nA new excel sheet titled 'Person Data' is addded to the 'Balance Sheet.xlsx' file with the transactions requested within selected dates"
        elif T != "":
            message += "\nA new excel sheet titled 'Tag Data' is addded to the 'Balance Sheet.xlsx' file with the transactions under this tag within selected dates"
        if str(self.Account.get()) != "":
            message += "\nA new excel sheet titled 'Account Data' is addded to the 'Balance Sheet.xlsx' file with the transactions involving this Account within Selected Dates"

        if str(self.Direction.get()) != "":

            if str(self.Account.get()) != "":
                    message += "\nA new excel sheet titled '(Account + Direction) Data' is addded to the 'Balance Sheet.xlsx' file with the transactions in the Selected Direction involving this Account within Selected Dates"

            else:
                message += "\nA new excel sheet titled 'Direction Data' is addded to the 'Balance Sheet.xlsx' file with the transactions in the Selected Direction  within Selected Dates"

        messagebox.showinfo("Results are here.....", message)


        #clearing the user input section after calculation and reult output to excel
        self.I.delete(0,"end")
        self.Tag.set("")
        self.Direction.set("")
        self.Account.set("")
        self.personVar.set(0)

        return

    #when the user cllicks on See Statistics, we will use this function to display more options regarding the user's interest for the result type he/she wants to see
    def Statistics_Options(self):

        #designing a frame for calendar and putting it up in the StatisticsFrame window
        self.startingCalendarFrame = LabelFrame(self.StatisticsFrame, borderwidth = 5, pady = 5, width = 315, height = 210, bg = 'white')
        self.startingCalendarFrame.grid(row = 1, column = 0, padx = 15, pady = 12, columnspan = 2, sticky = W)

        #getting current date from interacting with the computer for setting it as the default when the calendar is called for taking the user input
        today = str(_datetime.date.today())
        today = datetime.strptime(today, "%Y-%m-%d")

        #putting the calendar up inside the created frame startingCalendarFrame
        global startCal
        self.startCal = Calendar(self.startingCalendarFrame, selectmode = "day", year = today.year, day = today.day-7, month = today.month)
        self.startCal.grid(row = 0, column = 0, sticky = W, padx = 6, ipadx = 20, ipady = 0, columnspan = 2)

        # creating label widgets for user understanding what to enter in the box next to it and storing those inside a variable
        self.ToText = Label(self.StatisticsFrame, text = "- To -  ", bg = 'white', font = self.fontStyle).grid(row = 1, column = 2, columnspan = 1, sticky = W)

        #designing a frame for an inside calendar and putting it up in the StatisticsFrame window
        self.endingCalendarFrame = LabelFrame(self.StatisticsFrame, borderwidth = 5, pady = 5, width = 309, height = 210, bg = 'white')
        self.endingCalendarFrame.grid(row = 1, column = 3, pady = 12, sticky = W, columnspan = 3)
        self.endingCalendarFrame.grid_propagate(False)

        #putting the calendar up inside the created frame endingCalendarFrame
        global endCal
        self.endCal = Calendar(self.endingCalendarFrame, selectmode = "day", year = today.year, day = today.day, month = today.month)
        self.endCal.grid(row = 0, column = 0, sticky = E, padx = 10, ipadx = 12, ipady = 0)
        #self.endCal.grid_propagate(False)

        # use will input some item name / person name inside this text box
        self.ItemText = Label(self.StatisticsFrame, text = "Item/Person Name:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 2, column = 0, sticky = E)
        global I
        self.I = Entry(self.StatisticsFrame, width = 40,  borderwidth = 3, background = "white", foreground = "black", font = self.fontStyle)
        self.I.grid(row = 2, column = 1, columnspan = 4, sticky = W, padx = 0, pady = 21)

        # creating a checkbox for person/non-living thing as given and taken are relevant for persons and total is important for data related to non-living things
        global personVar
        self.personVar = IntVar()
        self.personCheck = Checkbutton(self.StatisticsFrame, text = 'Person', variable = self.personVar, background = 'white', font = self.fontStyle).grid(row = 2, column = 4, sticky = E, padx = 10)

        # if the user is interested in seeing data related to some particular type of transaction differentiated by the tags
        self.TagText = Label(self.StatisticsFrame, text = "Transaction Tag:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 3, column = 0, sticky = E)
        global Tag, Tags
        self.Tag = StringVar(self.StatisticsFrame)
        self.Tag.set("")
        #reading user input tags from excel file created when user input the tags preferences the first time
        self.Tags = pd.read_excel("Balance Sheet.xlsx", sheet_name = "Tags", skiprows = 0, header = None)#squeeze = True, 
        self.Tags = self.Tags.values.tolist()

        self.TagDropdownMenu = OptionMenu(self.StatisticsFrame, self.Tag, *self.Tags)
        self.TagDropdownMenu.grid(row = 3, column = 1, columnspan = 4, sticky = W, padx = 0, ipadx = 104, pady = 21)
        self.TagDropdownMenu.configure(font = self.fontStyle)


        self.AccountText = Label(self.StatisticsFrame, text = "Account:", justify = LEFT, bg = 'white', font = self.fontStyle).grid(row = 4, column = 0, sticky = E)

        self.Accounts = pd.read_excel("Balance Sheet.xlsx", sheet_name = "Sheet", header = None).iloc[0]#squeeze = True, 
        self.Accounts = self.Accounts.values.tolist()
        del self.Accounts[ 0 : 8 ]
        del self.Accounts[ (len(self.Accounts)-2) : len(self.Accounts) ]
        global Account
        self.Account = StringVar(self.StatisticsFrame)
        self.Account.set("")

        self.AccountDropdownMenu = OptionMenu(self.StatisticsFrame, self.Account, *self.Accounts)
        self.AccountDropdownMenu.grid(row = 4, column = 1, columnspan = 4, sticky = W, padx = 0, ipadx = 104, pady = 21)
        self.AccountDropdownMenu.configure(font = self.fontStyle)

        self.DirectionText = Label(self.StatisticsFrame, text = "   Direction of Transaction:", bg = 'white', padx = 5, font = self.fontStyle).grid(row = 5, column = 0, sticky = E, padx = 0)

        global Direction
        self.Direction = StringVar(self.StatisticsFrame)
        self.Direction.set("")

        self.DirectionDropdownMenu = OptionMenu(self.StatisticsFrame, self.Direction, "In", "Out", "Conversion")
        self.DirectionDropdownMenu.grid(row = 5, column = 1, columnspan = 4, sticky = W, padx = 0, ipadx = 104, pady = 21)
        self.DirectionDropdownMenu.configure(font = self.fontStyle)

        # creating a submit button for taking input of the transaction with the user's permission
        self.SubmitButton = Button(self.StatisticsFrame, text = "Submit", padx = 292, width = 13, pady = 5, background = "green", foreground = "white", command = self.Statistics_Calculations, font = self.fontStyle)
        self.SubmitButton.grid(row = 6, column = 0, padx = 0, pady = 12, columnspan = 7)
        #self.SubmitButton.grid_propagate(False)


    def start(self):
        self.StatisticsFrame.mainloop()
