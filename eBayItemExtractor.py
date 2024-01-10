from __future__ import print_function
from openpyxl import load_workbook
from googleapiclient.discovery import build
from google.oauth2 import service_account
import csv
import pandas as pd

class ebay:

    parts_list = []
    file_name = r"C:\Users\govil\Downloads\eBay-OrdersReport-Jan-01-2024-10_23_24-0700-13137996720.csv"

    def to_pandas(self):
        # confirm csv file and save folder
        if not self.file_name.endswith(".csv"):
            self.file_name += '.csv'

        # parse csv for info
        with open(self.file_name, newline='') as csvfile: 
            reader = csv.reader(csvfile)
            for i, row in enumerate(reader):
                if "Sales Record Number" in row[0]:
                    self.header = i
                    break
        
        # create data frame
        self.df = pd.read_csv(
            filepath_or_buffer=self.file_name, 
            header=self.header,
            nrows=21
            )
        
        # clean df of empty rows
        self.df.dropna(how='all', inplace=True)

        # create new df of needed info
        self.df = self.df[["Sale Date","Item Title","Sold For","Quantity"]]
        self.df["Sold For"] = self.df["Sold For"].str.replace('$', '')
        self.df[["Sold For", "Quantity"]] = self.df[["Sold For", "Quantity"]].apply(pd.to_numeric, errors='coerce')
        
        pattern = r'\(G\d{1,3}\)'
        self.df["G"] = self.df["Item Title"].str.contains(pattern)
        pattern = r'\(R\d{1,3}\)'
        self.df["R"] = self.df["Item Title"].str.contains(pattern)

        self.df["G"].fillna(False, inplace=True)
        self.df["R"].fillna(False, inplace=True)

        for ind in self.df.index:
            if self.df["G"][ind] != True and self.df["R"][ind] != True:
                self.df.drop([ind], inplace=True)
        
        self.df["Sold For"] *= self.df["Quantity"]
        self.df["Sold For"] = round(self.df["Sold For"],2)
        self.df.drop(columns=["Quantity"], inplace=True)

        self.df['Gov'] = None
        self.df['Ro'] = None
        for ind in self.df.index:
            sold_for = self.df.at[ind, 'Sold For']  # Accessing the 'Sold For' column using .at
            if self.df.at[ind, "R"] == True:
                self.df.loc[ind, 'Gov'] = round(sold_for * 0.275, 2)  # Using .loc for assignment
                self.df.loc[ind, 'Ro'] = round(sold_for * 0.225, 2)
            if self.df.at[ind, "G"] == True:
                self.df.loc[ind, 'Gov'] = round(sold_for * 0.5, 2)
                self.df.loc[ind, 'Ro'] = 0

        self.df.drop(columns=["G","R"], inplace=True)

        # Convert DataFrame to a list
        self.parts_list = self.df.values.tolist()
        for row in self.parts_list:
            # print(row)
            pass


    def addParts(self,sheetName):
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SERVICE_ACCOUNT_FILE = "keysEbay.json"
        # SERVICE_ACCOUNT_FILE = "ebay-372416-fc8582c01c9f.json"

        creds = None
        creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        
        # The ID of spreadsheet.
        SAMPLE_SPREADSHEET_ID = '1pTfI0LxrtkoxVoAEU0NvG6vBmR493UB8rBzZNC0OP2g'
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                        range=f'{sheetName}!B1:B').execute()
        titles = result.get('values', [])
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                        range=f'{sheetName}!A1:A').execute()
        dates = result.get('values', [])


        descrips = [title[0] for title in titles]
        dateChecks = [date[0] for date in dates]

        allItems = []
        allItems = [
            item for item in reversed(self.parts_list)
            if not any(
                descrips[k] == item[1] and dateChecks[k] == item[0]
                for k in range(len(descrips))
            )
        ]
        # check current year

        all_items = [item for item in allItems if item[0][-1] == '4']
        allItems = all_items

        print(f"{len(allItems)} items added to sheet!")
        for item in allItems:
           print(item)
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=f'{sheetName}!A1',valueInputOption='USER_ENTERED',
            body = {'values':allItems}).execute()
        return request

    def ebayUpdater(self): 
        self.to_pandas()
        self.addParts('2024 Shop')
        return None

a=ebay()
a.ebayUpdater()

