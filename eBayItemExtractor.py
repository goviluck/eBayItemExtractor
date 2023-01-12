from __future__ import print_function
from openpyxl import load_workbook
from googleapiclient.discovery import build
from google.oauth2 import service_account

def getParts(path):
    wb = load_workbook(rf'{path}')
    ws = wb.active
    partsList = []
    for i in range (1,100):
        index = 'X' + str(i)
        partName = ws[index].value
        indexDate = 'AX' + str(i)
        saleDate = ws[indexDate].value
        if type(saleDate) == str:
            if saleDate[-1] == '3': #change year date here
                if type(partName) == str:
                    #print(saleDate,partName)
                    if partName[-6:-4] == '(R' or partName[-5:-3] == '(R':
                        #x = partName.rfind('(')
                        price = round(ws['AB'+str(i)].value,1)
                        info = [ws['AX'+str(i)].value,partName,price,round(price*.275,2),round(price*.225,2)]
                        partsList.append(info)
                    elif partName[-6:-4] == '(G' or partName[-5:-3] == '(G':
                        #x = partName.rfind('(')
                        price = round(ws['AB'+str(i)].value,1)
                        info = [ws['AX'+str(i)].value,partName,price,round(price*.5,2),0]
                        partsList.append(info)
    #print(partsList)
    return partsList

def addParts(sheetName,partsList):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = 'Ebay\keysEbay.json'
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
    for i in range(len(partsList)):
        j=0
        for title in descrips:
            if title == partsList[i][1]:
                j=1
        item = partsList[i]
        if j==0:
            allItems.append(item)
    allItems.reverse()
    request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
        range=f'{sheetName}!A1',valueInputOption='USER_ENTERED',
        body = {'values':allItems}).execute()
    return request

def ebayUpdater():
    partsList = getParts(r"C:\Users\govil\Downloads\eBay-OrdersReport-Jan-11-2023-19_30_59-0700-1175751756.xlsx")
    addParts('2023 Shop',partsList)
    return None
ebayUpdater()