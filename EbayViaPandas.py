from __future__ import print_function
from openpyxl import load_workbook
from googleapiclient.discovery import build
from google.oauth2 import service_account
import csv
import pandas as pd
from tkinter import Tk, filedialog, messagebox
import sys
from datetime import datetime
import os
import glob


# pd.set_option('future.no_silent_downcasting', True)

# Define a class to handle eBay data processing
class EbayDataProcessor:
    
    # List to store sold parts from the CSV file
    sold_parts_list = []
    
    def __init__(self):
        # self.csv_file_path = self.select_csv_file()
        self.csv_file_path = self.select_latest_csv()

    def select_csv_file(self):
        # Initialize tkinter and hide the main window
        root = Tk()
        root.withdraw()

        # Open the file dialog to select a CSV file
        file_path = filedialog.askopenfilename(
            title="Select eBay Orders CSV File",
            filetypes=[("CSV Files", "*.csv")],
            initialdir="C:/Users/govil/Downloads/"
        )

        # Check if a file was selected
        if not file_path:
            messagebox.showerror("Error", "No file selected. Exiting program.")
            sys.exit()

        # Return the selected file path
        return file_path
    
    # File name of the CSV file containing eBay orders
    # csv_file_path = r"C:\Users\govil\Downloads\eBay-OrdersReport-Jun-19-2024-09_14_29-0700-11167508318.csv"

    def select_latest_csv(self):
        downloads_folder = "C:/Users/govil/Downloads/"
        csv_files = glob.glob(os.path.join(downloads_folder, "*.csv"))

        if not csv_files:
            messagebox.showerror("Error", "No CSV files found in Downloads folder.")
            sys.exit()

        # Get the most recently modified CSV file
        latest_file = max(csv_files, key=os.path.getmtime)

        return latest_file
    
    def parse_csv_to_dataframe(self):
        # Check if the file is a CSV file, if not append '.csv'
        if not self.csv_file_path.endswith(".csv"):
            self.csv_file_path += '.csv'

        # Read the CSV file to find the header row
        with open(self.csv_file_path, newline='') as csvfile:
            reader = csv.reader(csvfile)
            for i, row in enumerate(reader):
                if "Sales Record Number" in row[0]:
                    self.header_row_index = i
                    break

        # Create a DataFrame from the CSV file starting from the header row
        self.dataframe = pd.read_csv(filepath_or_buffer=self.csv_file_path, header=self.header_row_index)
        
        # Remove any empty rows from the DataFrame
        self.dataframe.dropna(how='all', inplace=True)

        # Select only the necessary columns
        self.dataframe = self.dataframe[["Sale Date", "Item Title", "Sold For", "Quantity", "Total Price"]]
        
        


        # Remove dollar signs from the "Sold For" and "Total Price" column and convert it to numeric
        self.dataframe["Sold For"] = self.dataframe["Sold For"].str.replace('$', '')
        self.dataframe["Total Price"] = self.dataframe["Total Price"].str.replace('$', '')
        self.dataframe[["Sold For", "Total Price", "Quantity"]] = self.dataframe[["Sold For", "Total Price", "Quantity"]].apply(pd.to_numeric, errors='coerce')

        # remove if refunded
        self.dataframe = self.dataframe[self.dataframe["Total Price"] >= self.dataframe["Sold For"]]
        self.dataframe.drop(columns=["Total Price"], inplace=True)


        # Define patterns to identify "G" and "R" tags in the item titles
        pattern_g = r'\((G|GL)\d{1,3}\)'
        pattern_r = r'\(R\d{1,3}\)'

        # Check if item titles contain "G" or "R" tags
        self.dataframe["Has_G_Tag"] = self.dataframe["Item Title"].str.contains(pattern_g, regex=True)
        self.dataframe["Has_R_Tag"] = self.dataframe["Item Title"].str.contains(pattern_r, regex=True)

        # Fill NaN values in "Has_G_Tag" and "Has_R_Tag" columns with False, 
        # then convert the column type to the most appropriate type (in this case, boolean).
        # This ensures that the columns contain no missing values and are correctly typed as boolean for further processing.
        self.dataframe["Has_G_Tag"] = self.dataframe["Has_G_Tag"].fillna(False).infer_objects(copy=False)
        self.dataframe["Has_R_Tag"] = self.dataframe["Has_R_Tag"].fillna(False).infer_objects(copy=False)

        # Drop rows that do not contain "G" or "R" tags
        for ind in self.dataframe.index:
            if not self.dataframe.at[ind, "Has_G_Tag"] and not self.dataframe.at[ind, "Has_R_Tag"]:
                self.dataframe.drop([ind], inplace=True)

        # Calculate the total amount sold for each item
        self.dataframe["Total_Sold_For"] = self.dataframe["Sold For"] * self.dataframe["Quantity"]
        self.dataframe["Total_Sold_For"] = round(self.dataframe["Total_Sold_For"], 2)

        # Remove unecessary columns
        self.dataframe.drop(columns=["Quantity"], inplace=True)
        self.dataframe.drop(columns=["Sold For"], inplace=True)

        # Add columns for Gov and Ro calculations
        self.dataframe['Gov_Amount'] = None
        self.dataframe['Ro_Amount'] = None
        for ind in self.dataframe.index:
            total_sold_for = self.dataframe.at[ind, 'Total_Sold_For']
            if self.dataframe.at[ind, "Has_R_Tag"]:
                self.dataframe.loc[ind, 'Gov_Amount'] = round(total_sold_for * 0.275, 2)
                self.dataframe.loc[ind, 'Ro_Amount'] = round(total_sold_for * 0.225, 2)
            if self.dataframe.at[ind, "Has_G_Tag"]:
                self.dataframe.loc[ind, 'Gov_Amount'] = round(total_sold_for * 0.5, 2)
                self.dataframe.loc[ind, 'Ro_Amount'] = 0

        # Drop the "Has_G_Tag" and "Has_R_Tag" columns as they are no longer needed
        self.dataframe.drop(columns=["Has_G_Tag", "Has_R_Tag"], inplace=True)

        # Convert DataFrame to a list of values
        self.sold_parts_list = self.dataframe.values.tolist()
        for row in self.sold_parts_list:
            # Print the rows if needed
            # print(row)
            pass
        
        # Print the dataframe if needed
        # print(self.dataframe)
        return None


    def add_parts_to_sheet(self, sheet_name):
        # Define the scope and credentials for Google Sheets API
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        # SERVICE_ACCOUNT_FILE = "keysEbay.json"
        SERVICE_ACCOUNT_FILE = r"C:\Users\govil\OneDrive\Documents\Python projects\Ebay\keysEbay.json"
        
        # Load the credentials from the service account file
        creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        
        # The ID of the Google Sheets spreadsheet
        spreadsheet_id = '1pTfI0LxrtkoxVoAEU0NvG6vBmR493UB8rBzZNC0OP2g'
        service = build('sheets', 'v4', credentials=creds)
        
        # Fetch the existing data from the sheet to avoid duplicates
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=f'{sheet_name}!B1:B').execute()
        item_titles = result.get('values', [])
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=f'{sheet_name}!A1:A').execute()
        sale_dates = result.get('values', [])

        # Extract titles and dates for duplicate checking
        item_titles_list = [title[0] for title in item_titles]
        sale_dates_list = [date[0] for date in sale_dates]

        new_items = []

        # Check for new items to be added to the sheet
        for item in reversed(self.sold_parts_list):
            found = False
            if (item[0][-2:] != sheet_name[2:4]):
                # Filter items to only include those from the current year (ending in '5')
                found = True
                continue
            for k in range(len(item_titles_list)):
                if item_titles_list[k] == item[1] and sale_dates_list[k] == item[0]:
                    found = True
                    break
            if not found:
                new_items.append(item)

        gov_money_made = 0
        # Print the number of items to be added
        print(f"{len(new_items)} items added to sheet!")
        for item in new_items:
            print(item)
            gov_money_made += item[3] # add total gov money
        
        first_date_str = new_items[0][0]  # get first date
        # Convert the date string to a datetime object
        first_date = datetime.strptime(first_date_str, '%b-%d-%y')
        today = datetime.today()   # Get today's date
        days_diff = (today - first_date).days # Calculate the difference in days

        print(f"gov money made ${gov_money_made:,.2f} in {days_diff} days")
        print(f"gov money made per day: ${gov_money_made/days_diff:,.2f}")
        
        # Append new items to the Google Sheet
        request = sheet.values().append(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A1',
            valueInputOption='USER_ENTERED',
            body={'values': new_items}
        ).execute()
        return request

    def update_ebay_data(self):
        # Update eBay data and add new parts to the Google Sheet
        self.parse_csv_to_dataframe()
        self.add_parts_to_sheet('2025 Shop')
        return None

# Create an instance of the EbayDataProcessor class and run the updater method
ebay_data_processor = EbayDataProcessor()
ebay_data_processor.update_ebay_data()