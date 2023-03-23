# author: Ammar
# date: 23/03/2023
# description: update google sheet with updated info from ubuntu pc

from dmidecode import DMIDecode
import subprocess
import os
# from google import ServiceAccountCredentials
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv
load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
JSON_FILE = os.path.join(DATA_DIR, 'gsCreds.json')

dmi = DMIDecode()

# Path to your service account credentials JSON file
creds = Credentials.from_service_account_file(JSON_FILE)

# Build the Sheets API client
service = build('sheets', 'v4', credentials=creds)

# Set the spreadsheet ID and range
spreadsheet_id = os.getenv('SPREADSHEET_ID')

worksheet_name = os.getenv('SHEET_NAME')
maxColumn = os.getenv('MAX_COLUMN')
worksheet_range = f'{worksheet_name}!A1:{maxColumn}'
domain = os.getenv('DOMAIN')
defaultUser = os.getenv('DEFAULT_USER')

def getHostname():
    #ubuntu get hostname from /etc/hostname
    with open('/etc/hostname', 'r') as file:
        hostname = file.read()
        #strip trailing whitespace, and remove trailing whitespace
        hostname = hostname.strip()

        #if hostame has <hostname>.DOMAIN remove .DOMAIN
        if domain in hostname:
            hostname = hostname.replace(domain, '')


        
        return hostname

def getGpu():

    
    #ubuntu get gpu from lspci, #get GPU model and memory using bash, gpu=$(lspci | grep VGA | awk -F ":" '{print $3}')
    gpu = subprocess.check_output(['lspci | grep VGA | awk -F ":" \'{print $3}\''], shell=True)
    #format output gpu remove b' and ', and remove trailing whitespace
    gpu = gpu.decode('utf-8').replace('b\'', '').replace('\',', '').strip()
    #     # gpu = gpu.decode('utf-8').replace('b\'', '').replace('\'', '')

    # gpu = subprocess.check_output(['lspci', '-vnn', '-d', '10de:']).decode('utf-8')
    gpu = str(gpu)
    #remove ending and starting space in gpu
    gpu = gpu.strip()


    #create if else if gpu == NVIDIA Corporation Device 2206 (rev a1) then gpu = NVIDIA RTX 3080 10 GB
    if gpu == 'NVIDIA Corporation Device 2206 (rev a1)': 
        gpu = 'NVIDIA RTX 3080 10 GB'
    elif gpu == 'NVIDIA Corporation GM206 [GeForce GTX 960] (rev a1)':
        gpu = 'NVIDIA GTX 960 4 GB'
    elif gpu == 'NVIDIA Corporation GM107 [GeForce GTX 750 Ti] (rev a2)':
        gpu = 'NVIDIA GTX 750 Ti 2 GB'
    elif gpu == 'NVIDIA Corporation GP106 [GeForce GTX 1060 3GB] (rev a1)':
        gpu = 'NVIDIA GTX 1060 3 GB'
    elif gpu == 'NVIDIA Corporation GP106 [GeForce GTX 1060 6GB] (rev a1)':
        gpu = 'NVIDIA GTX 1060 6 GB'
    elif gpu == 'NVIDIA Corporation GK106 [GeForce GTX 660] (rev a1)':
        gpu = 'NVIDIA GTX 660 2 GB'
    elif gpu == 'NVIDIA Corporation Device 1f0a (rev a1)':
        gpu = 'NVIDIA GTX 1650 4 GB'
    elif gpu == 'NVIDIA Corporation Device 1f08 (rev a1)':
        gpu = 'NVIDIA GTX 2060 6 GB'
    elif gpu == 'NVIDIA Corporation Device 2188 (rev a1)':
        gpu = 'NVIDIA GTX 1650 4GB'
    elif gpu == 'NVIDIA Corporation Device 2489 (rev a1)':
        gpu = 'NVIDIA RTX 3060 Ti 8GB'
    elif gpu == 'NVIDIA Corporation GK107 [GeForce GTX 650] (rev a1)':
        gpu = 'NVIDIA GTX 650 1 GB'
    elif gpu == 'NVIDIA Corporation GM107 [GeForce GTX 750] (rev a2)':
        gpu = 'NVIDIA GTX 750 1 GB'
    elif gpu == 'NVIDIA Corporation GP106 [GeForce GTX 1060 6GB Rev. 2] (rev a1)':
        gpu = 'NVIDIA GTX 1060 6 GB'
    elif gpu == 'NVIDIA Corporation Device 2216 (rev a1)':
        gpu = 'NVIDIA RTX 3080 10 GB'

    
    return gpu

    #grep only gpu name from getGpu output

def getRam():
    # print('Manufacturer:\t', dmi.manufacturer())
    # print('Model:\t\t', dmi.model())
    # print('Total RAM:\t{} GB'.format(dmi.total_ram()))
    ram = dmi.total_ram()


    

    return ram
def getUser():
    #get user login this computer using command who
    outputWho = subprocess.check_output(['who'])
    outputWho = outputWho.decode('utf-8')
    # print(outputWho)
    # print(outputWho.split(' ')[0])
    user = outputWho.split(' ')[0]

    if user == 'root' or user == defaultUser:
        #return blank if user is root or defaultUser
        return ''
        
    return user



workstation = getHostname().lower().strip()
# Get the values from the worksheet
result = service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id,
    range=worksheet_range
).execute()
values = result.get('values', [])
header_row = values[0]


# Find the index  of all the header columns
header_indices = {}
for i, header in enumerate(header_row):
    header_indices[header] = i

# Find the row index that contains the workstation
row_index = None
for i, row in enumerate(values):
    # if row and row[header_indices['Workstation']] == getHostname():
    if row and row[header_indices['Workstation']].lower().strip() == workstation:
        #match not
        row_index = i
        print(row_index+1)
        break

# ##Old code before refactor
# if row_index is not None:
#     user = getUser()
#     ram = getRam()
#     gpu = getGpu()

    
    #update user to column with header 'User In Use'
    #if user is blank then do not update
    # if user != '':
    # update_range = f'{worksheet_name}!{chr(header_indices["User In Use"] + 65)}{row_index + 1}'
    # print(update_range)
    # body = {
    #     'values': [[user]]
    # }
    # result = service.spreadsheets().values().update(
    #     spreadsheetId=spreadsheet_id,
    #     range=update_range,
    #     valueInputOption='USER_ENTERED',
    #     body=body
    # ).execute()

#     #update ram to column with header 'RAM'
#     update_range = f'{worksheet_name}!{chr(header_indices["RAM (GB)"] + 65)}{row_index + 1}'
#     print(update_range)
#     body = {
#         'values': [[ram]]
#     }
#     result = service.spreadsheets().values().update(
#         spreadsheetId=spreadsheet_id,
#         range=update_range,
#         valueInputOption='USER_ENTERED',
#         body=body
#     ).execute()

#     #update gpu to column with header 'GPU'
#     update_range = f'{worksheet_name}!{chr(header_indices["GPU"] + 65)}{row_index + 1}'
#     print(update_range)
#     body = {
#         'values': [[gpu]]
#     }
#     result = service.spreadsheets().values().update(
#         spreadsheetId=spreadsheet_id,
#         range=update_range,
#         valueInputOption='USER_ENTERED',
#         body=body
#     ).execute()

#     print(f'Updated {workstation} with user {user} and ram {ram} and gpu {gpu}')

## After refactor
def update ():
    if row_index is None:
        return

    user = getUser()
    ram = getRam()
    gpu = getGpu()

    update_cells = {
        "User In Use": user,
        "RAM (GB)": ram,
        "GPU": gpu
    }
    #if user is blank then do not update user just update ram and gpu
    if user == '':
        update_cells.pop("User In Use")
        
    for header, value in update_cells.items():
        col_index = header_indices.get(header)
        if col_index is not None:
            update_range = f"{worksheet_name}!{chr(col_index + 65)}{row_index + 1}"
            body = {"values": [[value]]}
            result = service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=update_range,
                valueInputOption="USER_ENTERED",
                body=body,
            ).execute()
            # print(f"Updated {workstation} {header} with {value}")

    print(f"Updated {workstation} with user {user} and ram {ram} and gpu {gpu}")

update()