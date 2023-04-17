from __future__ import print_function
from openpyxl import load_workbook
from googleapiclient.discovery import build
from google.oauth2 import service_account

class ebay:

    partsList = []
    path = r"C:\Users\govil\Downloads\eBay-OrdersReport-Apr-17-2023-07_22_59-0700-12101679358.xlsx"

    def getParts(self):
        wb = load_workbook(rf'{self.path}') #load sheet
        ws = wb.active
        for i in range (1,100):
            index = 'X' + str(i)  #title index
            partName = ws[index].value  #title
            indexDate = 'AX'+str(i)  #date index
            saleDate = ws[indexDate].value  #sale date
            if type(saleDate) == str:  #if sale date exists
                if saleDate[-1] == '3': #change year date here
                    if type(partName) == str:  #if part exists
                        #print(saleDate,partName)
                        if partName[-6:-4] == '(R' or partName[-5:-3] == '(R':  #check if Ro part
                            #x = partName.rfind('(')
                            price = round((ws['AB'+str(i)].value)*(ws['AA'+str(i)].value),1) #multiply price and quanity
                            info = [saleDate,partName,price,round(price*.275,2),round(price*.225,2)] #collumns
                            self.partsList.append(info)
                        elif partName[-6:-4] == '(G' or partName[-5:-3] == '(G':
                            #x = partName.rfind('(')
                            price = round((ws['AB'+str(i)].value*(ws['AA'+str(i)].value)),1)
                            info = [saleDate,partName,price,round(price*.5,2),0]
                            self.partsList.append(info)

    def addParts(self,sheetName):
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        SERVICE_ACCOUNT_FILE = "OneDrive - Northeastern University\Ebay\keysEbay.json"
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
        descrips = []
        for i in range(len(titles)):
            descrips.append(titles[i][0])

        allItems = []
        for i in range(len(self.partsList)):
            j=0
            for title in descrips:
                if title == self.partsList[i][1]:
                    j=1
            item = self.partsList[i]
            if j==0:
                allItems.append(item)
        allItems.reverse()
        print(f"{len(allItems)} items added to sheet!")
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=f'{sheetName}!A1',valueInputOption='USER_ENTERED',
            body = {'values':allItems}).execute()
        return request

    def ebayUpdater(self): 
        self.getParts()
        self.addParts('2023 Shop')
        return None

a=ebay()
a.ebayUpdater()
