from Google import Create_Service
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

client = gspread.authorize(creds)

sheet = client.open("Certi_sheet").sheet1  # Open the spreadsheet

data = sheet.get_all_records()  # Get a list of all records

length = len(data)  # gives the number of items in the sheet


CLIENT_SECRET_FILE = 'client.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

folder_id = '1Ua3nGP8E2bPmfhFNapVtnuRhFKLRDque'
query = f"parents = '{folder_id}'"

response = service.files().list(q=query).execute()
files = response.get('files')
nextPageToken = response.get('nextPageToken')

while nextPageToken:
    response = service.files().list(q=query).execute()
    files.extend(response.get('files'))
    nextPageToken = response.get('nextPageToken')

pd.set_option('display.max_columns', 100)
pd.set_option('display.max_rows', 500)
pd.set_option('display.min_rows', 500)
pd.set_option('display.max_colwidth', 150)
pd.set_option('display.width', 200)
pd.set_option('expand_frame_repr', True)
df = pd.DataFrame(files)
link1 = "https://drive.google.com/uc?export=download&id="
print(df)
i = 0
j = 1
Column_B = "ThreeJS workshop"
Column_C = "3DA_"
Column_D = "certi-img/3DA_"

while j <= length:
    row = sheet.row_values(j + 1)
    name = row[4]
    print(name)
    name2 = name + ".png"
    sheet.update_cell(j+1,2, "ThreeJS workshop")
    sheet.update_cell(j + 1, 3, "3DA_"+name+".png")
    sheet.update_cell(j + 1, 4, "certi-image/3DA_" + name + ".png")
    while i<=29:
        value2 = df.iloc[i]["name"]
        if(name2 == value2):
            value = df.iloc[i]["id"]
            link = link1 + value
            print(link)
            sheet.update_cell(j + 1, 7, link)
        i = i + 1
    #sheet.update_cell(j + 1, 7, link)
    sheet.update_cell(j + 1, 8, "https://isavesit.org.in/isa-app/certificate/certi-img/WAW_" + name + ".png")
    time.sleep(20)
    i = 0
    j = j + 1