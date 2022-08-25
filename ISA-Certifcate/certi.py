from Google import Create_Service
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("certi.json", scope)

client = gspread.authorize(creds)

# sheet = client.open("GraphUp_with_Chart.Js_Certis").sheet1  # Open the spreadsheet
# sheet = client.open("Elucimate_Certis").sheet1
# sheet = client.open("3D_Printing_Certis").sheet1
# sheet = client.open("Tkinter_Certi").sheet1
# sheet = client.open("PP").sheet1
sheet = client.open("IoT_WS_Certi_2022").sheet1
data = sheet.get_all_records("IoT_WS_Certi_2022")  # Get a list of all records

length = len(data)  # gives the number of items in the sheet

print(length)

# Google drive wala part

CLIENT_SECRET_FILE = 'client_secret_file.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
# folder_id = '1Ks7q_uxPYw5z7RGOIV85tA-jd8-9-xZn' #- chart.js
# folder_id = '1gve5rOvYeWUu15BK_zNXx_eTfufxXKED' #- elucimate 2.0
# folder_id = '11-4plzeHg-nELtHmzXxBSKgdnRi4ZEwo'   #- 3D printing
# folder_id = '1ZG4CCbmkK52fkXA49Dr6C4ik0OGylqPB' #- Winners
# folder_id = '18lhrN9iLaZL7Qi02MPFieX_cMqo1N7iH' #- Tkinter
# folder_id = '1J6tQ0nt9DRPAtrmUcA2N7Hfzx8Fi09HQ' # praxis participation
folder_id = '1UFv8GyQ0y51f48OziR941cYcA-uu2hiK' # Iot Workshop 2022
query = f"parents = '{folder_id}'"
print('first')
response = service.files().list(q=query).execute()
files = response.get('files')
nextPageToken = response.get('nextPageToken')
print('second')
while nextPageToken:
    response = service.files().list(q=query).execute()
    files.extend(response.get('files'))
    nextPageToken = response.get('nextPageToken')
print('third')
df = pd.DataFrame(files)
# print(df)
pd.set_option('display.max_columns', 100)
pd.set_option('display.max_rows', 500)
pd.set_option('display.min_rows', 500)
pd.set_option('display.max_colwidth', 150)
pd.set_option('display.width', 200)
pd.set_option('expand_frame_repr', True)
df = pd.DataFrame(files)
print('fourth')
link1 = "https://drive.google.com/uc?export=download&id="
print(link1)
print(df)

i = 0
j = 0
while j <= length:
    row = sheet.row_values(j + 2)
    _id = row[0]
    print('first step - ', _id)
    # name = row[5]
    _id = _id+'.jpg'
    print(_id)
    while i<=18:
        value2 = df.iloc[i]["name"]
        if(_id == value2):
            value_id = df.iloc[i]["id"]
            link = link1 + value_id + ' '
            print(link)
            sheet.update_cell(j + 2, 7, link)
        i = i + 1
    sheet.update_cell(j + 2, 8, "https://isavesit.org.in/isa-app/certificate/certi-img/" + _id)
    print("https://isavesit.org.in/isa-app/certificate/certi-img/" + _id)
    time.sleep(20)
    i = 0
    j = j + 1

# # print(dir(service))
# print('done')