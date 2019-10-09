from __future__ import print_function
import pickle
import os.path
import sys
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import argparse

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive.metadata.readonly']

parser = argparse.ArgumentParser(description='sheets-to-html-gallery')
parser.add_argument('-l','--loglevel', default="WARNING", help='Log-level, e.g. WARNING, INFO, DEBUG', required=False)
parser.add_argument('-f','--file_id', help='ID of sheet from Google', required=True)
# how to find the Team Drive ID:
#     https://stackoverflow.com/questions/54300175/how-to-get-teamdriveid-when-use-google-drive-api
parser.add_argument('-t','--team_drive_id', help='ID of team drive from Google', required=True)
args = parser.parse_args()

import logging
# logging.basicConfig(level=getattr(logging, args.loglevel))

OBJECT_HEADER_RANGE = 'Inventory!A1:Z1'
# Note: number of objects in the sheet is hardcoded here.  Should be replaced with a sheets query
OBJECT_SHEET_RANGE = 'Inventory!A2:Z800'
COLUMN_NAME_ID = 'ID'
COLUMN_NAME_TITLE = 'Title'
COLUMN_NAME_LOCATION = 'Location'

#category_dict = {"oid00":{"count":0,"type":"fine_arts"}}
category_type_list = ["fine_arts", "silver", "ceramics", "glass", "metals", "furniture5", "furniture6", \
                 "textiles7", "textiles8", "textiles9", "accessories", "adornments", "doc_artifacts", \
                 "needlework", "books", "not_in_collection", "on_loan"]
category_dict = {}
for i, type in enumerate(category_type_list):
    prefix = "oid" + format(i, "02d")
    if i == 14: prefix = "oid20"
    if i == 15: prefix = "oid90"
    if i == 16: prefix = "oid98"
    category_dict[prefix] = {"pic_count":0,"type":type,"obj_id_count":0,"wo_pic_list":[], \
                             "object_count":0, "object_sets": {}}

# object_id_count is the number of objects without the sets expanded, \
#        where sets are with all the objects in the set being the same
# object_count is the number of objects with the sets expanded

service = None

def main():
    global service
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    obj_id_count = 0
    objects_with_pic_count = 0
    objects_without_pic_count = 0
    objects_unassigned = 0
    objects_deaccessioned = 0

    previous_set_list = []

    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # Build object header dictionary
    result = sheet.values().get(spreadsheetId=args.file_id, range=OBJECT_HEADER_RANGE).execute()
    col_list = result.get('values')[0]
    obj_col_list = {col_list[i]: i for i in range(0, len(col_list))}
    # print(obj_col_list)

    result = sheet.values().get(spreadsheetId=args.file_id, range=OBJECT_SHEET_RANGE).execute()
    object_array = result.get('values', [])
    if not object_array or len(object_array) < 1:
        print("Error: no data fetched from " + OBJECT_SHEET_RANGE)
        sys.exit(1)

    # after getting the rows from the spreadsheet, then prepare to do drive queries to get the pictures google drive ID
    service = build('drive', 'v3', credentials=creds)
    query_prefix = "trashed = false and not name contains 'detail' and name contains "

    for obj_row in object_array:
        if len(obj_row) > obj_col_list[COLUMN_NAME_LOCATION]:
            if "unassigned" in obj_row[obj_col_list[COLUMN_NAME_LOCATION]].lower():
                objects_unassigned += 1;
                # print("Unassigned object: {}".format(obj_row[obj_col_list[COLUMN_NAME_ID]]))
                continue
            if "deaccessioned" in obj_row[obj_col_list[COLUMN_NAME_LOCATION]].lower():
                objects_deaccessioned += 1;
                # print("Deaccessioned object: {}".format(obj_row[obj_col_list[COLUMN_NAME_ID]]))
                continue
        # else:
        #     print("No location for: {}".format(obj_row[obj_col_list[COLUMN_NAME_ID]]))

        obj_id_count += 1
        # if obj_id_count > 20:
        #     break
        # extract the object ID (without the oid prefix and the trailing info after the first dash
        object_id = obj_row[obj_col_list[COLUMN_NAME_ID]].split('-')[0][3:].lstrip("0")
        if object_id not in previous_set_list:
            obj_prefix = obj_row[obj_col_list[COLUMN_NAME_ID]][:5]
            if obj_prefix not in category_dict:
                print("Warning: bogus prefix for object: " + obj_row[obj_col_list[COLUMN_NAME_ID]])
                exit()
            category_dict[obj_prefix]['obj_id_count'] += 1
            category_dict[obj_prefix]['object_count'] += 1

            # determine if the object is a matching set, in order to count how many objects are in the set
            # and to NOT require a picture per object
            id_parts = object_id.split('_')
            if len(id_parts) > 2:
                previous_set_list = []
                set_size = 0
                for c in range(ord(id_parts[1]), ord(id_parts[2]) + 1):
                    previous_set_list.append(id_parts[0] + "_" + chr(c))
                    set_size += 1
                # increment object_count by set_size -1 because it was already incremented above
                category_dict[obj_prefix]['object_count'] += (set_size - 1)
                category_dict[obj_prefix]['object_sets'][object_id] = set_size

            # query to get picture file ID
            query = query_prefix + "'" + obj_row[obj_col_list[COLUMN_NAME_ID]] + "'"
            results = service.files().list(
                fields="files(name, id)",
                corpora='teamDrive', includeTeamDriveItems='true', supportsTeamDrives='true',
                teamDriveId=args.team_drive_id,
                q=query).execute()
            files_dict = results.get('files', [])
            if not files_dict or len(files_dict) < 1 and object_id not in previous_set_list:
                # if len(previous_set_list) > 0:
                #     print("{} -- previous list: ".format(obj_row[obj_col_list[COLUMN_NAME_ID]]), end =" ")
                #     print(*previous_set_list, sep=", ")
                objects_without_pic_count += 1
                category_dict[obj_prefix]['wo_pic_list'].append(object_id)

                # print("Warning: no pic for object: {:12} {:64.64}    Location: {}". \
                #       format(obj_row[obj_col_list[COLUMN_NAME_ID]], \
                #              obj_row[1], \
                #              obj_row[obj_col_list[COLUMN_NAME_LOCATION]]))
            else:
                if len(files_dict) > 1:
                    print("Warning: multiple pics for object: " + obj_row[obj_col_list[COLUMN_NAME_ID]])
                objects_with_pic_count += 1
                category_dict[obj_prefix]['pic_count'] += 1

                #print(files_dict)
                #print(files_dict[0]["name"][:5])
                # pic_id = next(iter(files_dict))['id']
                # # print(u'{0}: {1}'.format(row[obj_col_list[COLUMN_NAME_ID]], pic_id))
                # path = "https://drive.google.com/a/sargenthouse.org/thumbnail?id=" + pic_id

    # write out report file
    import time
    report_filename = time.strftime("%Y-%m-%d") + "_object-report.csv"
    report_file = open(report_filename, "w")
    report_file.write("Range,Category,Object_ID_Count,Set_Count,Object_Count,Pic_Count,WO_Pics,WO_Pics_List\n")
    for category in category_dict:
        objs_wo_pics = category_dict[category]['obj_id_count'] - category_dict[category]['pic_count']
        objs_wo_pics_list = '"' + ', '.join(category_dict[category]['wo_pic_list']) + '"'
        report_file.write("{}00,{},{},{},{},{},{},{}\n".format(category[3:5],category_dict[category]['type'], \
                                                                      category_dict[category]['obj_id_count'], \
                                                                      len(category_dict[category]['object_sets']), \
                                                                      category_dict[category]['object_count'], \
                                                                      category_dict[category]['pic_count'], \
                                                                      objs_wo_pics, objs_wo_pics_list))
    report_file.write(",,,,,\n")
    report_file.write("Any.,totals,{0},,,{1},{2},\n".format(obj_id_count, \
                                                            objects_with_pic_count, objects_without_pic_count))

    print("Unassigned objects: {}  -- deaccessioned objects: {} ".format(objects_unassigned, objects_deaccessioned))


if __name__ == '__main__':
    main()
