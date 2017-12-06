import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name('My Project 65235-f09324c83f8c.json', scope)

gc = gspread.authorize(credentials)

spreadsheet = gc.open("Copy of composite V2.xlsm")
worksheet = spreadsheet.worksheet("Test Sheet")

val = worksheet.acell('E12').value
# print(val)

data_object = list()

row = 2
list_of_lists = worksheet.get_all_values()
# print(list_of_lists)
for grantee in list_of_lists:
    grantee_object = dict()
    if grantee[0] != "Grantee":
        grantee_object["project_name"] = grantee[1]
        grantee_object["country"] = grantee[2]
        grantee_object["region"] = grantee[3]
        grantee_object["duration"] = grantee[4]
        grantee_object["project_type"] = grantee[5]
        grantee_object["project_description"] = grantee[6]
        grantee_object["monthly_cost"] = grantee[7]
        data_object.append(grantee_object)
        print(grantee_object)
# print(data_object)
with open('final_data.json', 'w') as outfile:
    json.dump(data_object, outfile)
