from __future__ import print_function
import pickle
import os.path
import sys
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from dominate import document
from dominate.tags import *
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
logging.basicConfig(level=getattr(logging, args.loglevel))

OBJECT_HEADER_RANGE = 'Inventory!A1:Z1'
OBJECT_SHEET_RANGE = 'Inventory!A2:Z800'
COLUMN_NAME_LOCATION = 'Location'
COLUMN_NAME_ID = 'ID'
COLUMN_NAME_TITLE = 'Title'
COLUMN_NAME_OBJECT_TYPE = 'Object_Type'
COLUMN_NAME_CREATOR = 'Creator'
COLUMN_NAME_CREATION_DATE = 'Creation_Date'
COLUMN_NAME_SUBJECT_STYLE = 'Subject_Style'
COLUMN_NAME_BINDER_DESC = 'Binder_Description'
COLUMN_NAME_MEDIUM = 'Medium'
PEOPLE_HEADER_RANGE = 'People!A1:Z1'
PEOPLE_SHEET_RANGE = 'People!A2:Z100'
TYPE_PORTRAITS = "portrait"
TYPE_SILHOUTTE = "silhouette"
PEOPLE_COL_NAME = "Full_Name"
PEOPLE_COL_REL_TO_JUDITH = "RelationshipToJudith"
PEOPLE_COL_DESCRIP = "Description"
PEOPLE_COL_URL = "URL"

service = None

def main(location='Best Parlor'):
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

    # Build people header dictionary
    result = sheet.values().get(spreadsheetId=args.file_id, range=PEOPLE_HEADER_RANGE).execute()
    col_list = result.get('values')[0]
    people_col_list = {col_list[i]: i for i in range(0, len(col_list))}
    # print(people_col_list)

    result = sheet.values().get(spreadsheetId=args.file_id, range=PEOPLE_SHEET_RANGE).execute()
    people_array = result.get('values', [])
    if not people_array or len(people_array) < 1:
        print("Error: no data fetched from " + PEOPLE_SHEET_RANGE)
        sys.exit(1)

    # after getting the rows from the spreadsheet, then prepare to do drive queries to get the pictures google drive ID
    service = build('drive', 'v3', credentials=creds)
    query_prefix = "trashed = false and not name contains 'detail' and name contains "

    doc = document(title=location)
    with doc.head:
        link(rel='stylesheet', href='shm-binder.css')
        script(type='text/javascript', src='shm-binder.js')
        meta(name="viewport", content="width=device-width, initial-scale=1")
    with doc.body:
        # < div id = "newsbar" class ="page_title" > < / div >
        div(_class ="page_title").add(location)
        # div(style="text-align:center").add(h2(location))
        span(_class="popuptext", id="myPopup")

    for obj_row in object_array:
        person_row = None
        creator_row = None
        style_row = None

        # check row length to verify the location column is included
        if len(obj_row) > obj_col_list[COLUMN_NAME_LOCATION] and location in obj_row[obj_col_list[COLUMN_NAME_LOCATION]]:
            # print(obj_row)
            # get picture ID for object to use as img src
            query = query_prefix + "'" + obj_row[obj_col_list[COLUMN_NAME_ID]] + "'"
            results = service.files().list(
                fields = "files(name, id)",
                corpora='teamDrive', includeTeamDriveItems='true', supportsTeamDrives='true',
                teamDriveId = args.team_drive_id,
                q = query).execute()
            files_dict = results.get('files', [])

            if not files_dict or len(files_dict) < 1:
                print("Warning: no pic for object: " + obj_row[obj_col_list[COLUMN_NAME_ID]] + " " + obj_row[1] + " so skipping it")
            else:
                if len(files_dict) > 1:
                    print("Warning: multiple pics for object: " + obj_row[obj_col_list[COLUMN_NAME_ID]])
                pic_id = next(iter(files_dict))['id']
                # print(u'{0}: {1}'.format(row[obj_col_list[COLUMN_NAME_ID]], pic_id))
                path = "https://drive.google.com/a/sargenthouse.org/thumbnail?id="+pic_id

                # make the pop-up text
                alt_text = ""
                if len(obj_row) > obj_col_list[COLUMN_NAME_OBJECT_TYPE]:
                    # print(u'{0}: {1}'.format(row[obj_col_list[COLUMN_NAME_ID]], row[obj_col_list[COLUMN_NAME_OBJECT_TYPE]]))

                    if TYPE_PORTRAITS in obj_row[obj_col_list[COLUMN_NAME_OBJECT_TYPE]] or \
                            TYPE_SILHOUTTE in obj_row[obj_col_list[COLUMN_NAME_OBJECT_TYPE]]:
                        # add subject, start by gettng data no the person from the people tab
                        for tmp_row in people_array:
                            object_subject = obj_row[obj_col_list[COLUMN_NAME_SUBJECT_STYLE]].lower().strip()
                            person_name = tmp_row[people_col_list[PEOPLE_COL_NAME]].lower().strip()
                            # print('Comparing: ' + object_subject + ' to: ' + person_name)
                            if object_subject == person_name:
                                person_row = tmp_row
                                break
                        if person_row is None:
                            print("No data for: " + obj_row[obj_col_list[COLUMN_NAME_SUBJECT_STYLE]])
                        else:
                            subject = reverse_name(obj_row[obj_col_list[COLUMN_NAME_SUBJECT_STYLE]])[0]
                            if has_data(person_row, people_col_list[PEOPLE_COL_URL]):
                                # add the subject's name as a link
                                alt_text += "<a target='_blank' href='" + person_row[people_col_list[PEOPLE_COL_URL]] + \
                                            "'>" + subject + '</a>'
                            else:
                                alt_text += '<b>' + subject + '</b>'
                            # add birth/death dates
                            # if has_data(person_row, people_col_list[PEOPLE_COL_BIRTH]) and \
                            #         has_data(person_row, people_col_list[PEOPLE_COL_DEATH]):
                            #     alt_text += ' (' + person_row[people_col_list[PEOPLE_COL_BIRTH]]
                            #     alt_text += '-' + person_row[people_col_list[PEOPLE_COL_DEATH]] + ')'
                            if has_data(person_row, people_col_list[PEOPLE_COL_REL_TO_JUDITH]):
                                alt_text += " <i>Judith's " + person_row[
                                    people_col_list[PEOPLE_COL_REL_TO_JUDITH]] + "</i>"
                        alt_text += '<br>'
                    else:
                        # add style & description
                        if has_data(obj_row, obj_col_list[COLUMN_NAME_SUBJECT_STYLE]):
                            for tmp_row in people_array:
                                object_style = obj_row[obj_col_list[COLUMN_NAME_SUBJECT_STYLE]].lower().strip()
                                style_name = tmp_row[people_col_list[PEOPLE_COL_NAME]].lower().strip()
                                # print('Comparing: ' + object_subject + ' to: ' + person_name)
                                if object_style in style_name:
                                    style_row = tmp_row
                                    break
                            if style_row is None:
                                print("No data for style: " + obj_row[obj_col_list[COLUMN_NAME_SUBJECT_STYLE]])
                            else:
                                if has_data(style_row, people_col_list[PEOPLE_COL_URL]):
                                    alt_text += "<a target='_blank' href='" + style_row[
                                        people_col_list[PEOPLE_COL_URL]] + \
                                                "'>" + obj_row[obj_col_list[COLUMN_NAME_SUBJECT_STYLE]] + '</a>'
                                else:
                                    alt_text += obj_row[obj_col_list[COLUMN_NAME_SUBJECT_STYLE]]
                                alt_text += ' style<br>'
                        if has_data(obj_row, obj_col_list[COLUMN_NAME_BINDER_DESC]):
                            alt_text += obj_row[obj_col_list[COLUMN_NAME_BINDER_DESC]] + '<br>'

                    # Add creator:
                    if has_data(obj_row, obj_col_list[COLUMN_NAME_CREATOR]) and \
                            'unknown' not in obj_row[obj_col_list[COLUMN_NAME_CREATOR]].lower():
                        # look for data on the creator
                        for tmp_row in people_array:
                            creator = obj_row[obj_col_list[COLUMN_NAME_CREATOR]].lower().strip()
                            person_name = tmp_row[people_col_list[PEOPLE_COL_NAME]].lower().strip()
                            # print('Comparing: ' + creator + ' to: ' + person_name)
                            # note: the creator may have trailing data, e.g. Attributed to
                            if person_name in creator:
                                creator_row = tmp_row
                                break
                        if creator_row is None:
                            print("No data for: " + obj_row[obj_col_list[COLUMN_NAME_CREATOR]])
                        creator_plus_attribute = reverse_name(obj_row[obj_col_list[COLUMN_NAME_CREATOR]])
                        # creator = reverse_name(obj_row[obj_col_list[COLUMN_NAME_CREATOR]])[0]
                        if creator_plus_attribute[1] is None:
                            alt_text += 'by '
                        else:
                            alt_text += creator_plus_attribute[1] + ' '
                        # <a target='_blank' href='https://en.wikipedia.org/wiki/James_Frothingham'>James Frothingham</a>
                        if has_data(creator_row, people_col_list[PEOPLE_COL_URL]):
                            alt_text += "<a target='_blank' href='" + creator_row[people_col_list[PEOPLE_COL_URL]] + \
                                        "'>" + creator_plus_attribute[0] + '</a>'
                        else:
                            alt_text += creator_plus_attribute[0]
                    # Add creation date
                    if has_data(obj_row, obj_col_list[COLUMN_NAME_CREATION_DATE]):
                        alt_text += ' (' + obj_row[obj_col_list[COLUMN_NAME_CREATION_DATE]] + ')'

                    if has_data(obj_row, obj_col_list[COLUMN_NAME_MEDIUM]):
                        alt_text += ' ' + obj_row[obj_col_list[COLUMN_NAME_MEDIUM]] + '.'
                    alt_text += '<br>'

                    if has_data(person_row, people_col_list[PEOPLE_COL_DESCRIP]):
                        alt_text += '<div style="text-align:left">' + \
                                    person_row[people_col_list[PEOPLE_COL_DESCRIP]] + '</div>'

                doc.body += div(img(src=path, alt=alt_text, style="width:100%", onclick="myFunctionPopUp(this);"), _class='column')

    # print(doc)
    out_file = location.replace(" ", "_") + '.html'
    with open(out_file, 'w') as f:
        f.write(doc.render())

def has_data(row, column):
    if row is not None and len(row) > column and len(row[column]) > 0:
        return True
    else:
        return False

def has_match(row, column, data):
    if row is not None and len(row) > column and data in row[column] :
        return True
    else:
        return False

def reverse_name(name):
    import re
    # remove date in parathesis.  Dates can be of the form 'died 1900'
    # print("processing name: " + name)
    match = re.search("\(.*[0-9]\)", name)
    if match is not None:
        # print("match start:" + str(match.start()) + "match end: " + str(match.end()))
        name_wo_date = name[0:match.start()] + name[match.end():]
    else:
        name_wo_date = name
    last_first = name_wo_date.split(",")
    if len(last_first) != 2:
        print("The following is not in the last, first name format:" + name)
        result = None, None
    else:
        # deal with stuff in parens after the first name after the date has been removed
        first_and_parens = last_first[1].split("(")
        if len(first_and_parens) == 1:
            result = last_first[1].lstrip() + last_first[0].lstrip(), None
        elif len(first_and_parens) == 2:
            # print(last_first[0] + " " + last_first[1] + " -- " + first_and_parens[0] + " " + first_and_parens[1])
            text_after_date = first_and_parens[1].lstrip().split(")")[0]
            result = first_and_parens[0].lstrip() + last_first[0].lstrip(), text_after_date
            # print(result)
        else:
            print("The following has data in multiple parens:" + name)
            result = None, None
    return result

if __name__ == '__main__':
    room_list = ['Murray Room', 'Best Parlor', 'JSSargent Room', 'Judiths Room']
    room_list = ['Murray Room']
    # room_list = ['JSSargent Room']
    # room_list = ['Judiths Room']
    for room in room_list:
        main(room)



#------------- test of batchGetByDataFilter
# filter_request_body = {
#     'data_filters': [ { "a1Range": SHEET_RANGE } ]
# }

# result = sheet.values().batchGetByDataFilter(spreadsheetId=args.file_id,
#                                         body=filter_request_body).execute()


#------------- list basic filters:
# def get_existing_basic_filters(wkbkId: str) -> dict:
#     global service
#     params = {'spreadsheetId': wkbkId,
#               'fields': 'sheets(properties(sheetId,title),basicFilter)'}
#     response = service.spreadsheets().get(**params).execute()
#     # Create a sheetId-indexed dict from the result
#     filters = {}
#     for sheet in response['sheets']:
#         if 'basicFilter' in sheet:
#             filters[sheet['properties']['sheetId']] = sheet['basicFilter']
#     return filters


    # my_filters = get_existing_basic_filters(args.file_id)
    # print(my_filters)
