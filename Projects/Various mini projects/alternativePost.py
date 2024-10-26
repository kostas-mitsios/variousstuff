import os
import json
import httplib2
from googleapiclient.discovery import build
from google.oauth2 import service_account

credentials = service_account.Credentials.from_service_account_file('C:/Users/katikako/Downloads/credentials.json')
drive_service = build('drive', 'v3', credentials=credentials)
sheets_service = build('sheets', 'v4', credentials=credentials)

file_id = '1sm_ePXHiTxty5gHfgW_rWGhzdnml1G6sVfuwKfKWeMM'  # Replace with the actual file ID of your Google Sheet

# Search for a named range
named_range_name = 'DataSet'  # Replace with the actual named range you want to search for
named_ranges = sheets_service.spreadsheets().get(spreadsheetId=file_id).execute()['namedRanges']
named_range = next((nr for nr in named_ranges if nr['name'] == named_range_name), None)

print(named_ranges)

if named_range:

    start_row_index = named_range['range']['startRowIndex']
    end_row_index = named_range['range']['endRowIndex']

    print(start_row_index, end_row_index)

    # Increase the end row index by 1
    new_end_row_index = end_row_index + 1
    print(new_end_row_index)

    # Calculate the range of the named range
    #sheet_id = named_ranges['sheets'][0]['properties']['sheetId']
    sheet_name = 'Dogs'
    named_range_range = f"{sheet_name}!A{start_row_index+1}:C{new_end_row_index}"
    print(named_range_range)

    # Update the named range using values update
    body = {
        'values': [[''] * (named_range['range']['endColumnIndex'] + 1)]  # Insert an empty row
    }

    for sheet in named_ranges['sheets']:
        if sheet['properties']['title'] == 'Sheet1':
            sheet_id = sheet['properties']['sheetId']
            break

    sheet_properties = sheets_service.spreadsheets().get(
        spreadsheetId=file_id,
        ranges=f'Sheet1!A1:C5',  # Modify the range as needed
        fields='sheets(properties.gridProperties)'
    ).execute()

    current_grid_properties = sheet_properties['sheets'][0]['properties']['gridProperties']

    # Increase the row count by 1
    updated_row_count = current_grid_properties['rowCount'] + 1

    # Update the grid properties of "Sheet1"
    requests = [
        {
            'updateSheetProperties': {
                'properties': {
                    'sheetId': sheet_id,
                    'gridProperties': {
                        'rowCount': updated_row_count,
                        'columnCount': current_grid_properties['columnCount']
                    }
                },
                'fields': 'gridProperties(rowCount)'
            }
        }
    ]

    # Execute the requests to update the grid properties
    request = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=file_id,
        body={'requests': requests}
    )
    response = request.execute()
"""
    request = sheets_service.spreadsheets().values().update(
        spreadsheetId=file_id,
        range=named_range_range,
        valueInputOption='RAW',
        body=body
    )
    response = request.execute()
    print('Named range updated successfully!')
else:
    print(f'Named range "{named_range_name}" not found.')


values = [
    ['A1', 'B1', 'C1'],  # Row 1
    ['A2', 'B2', 'C2'],  # Row 2
    ['A3', 'B3', 'C3'],  # Row 3
]

body = {'values': values}
request = sheets_service.spreadsheets().values().update(
    spreadsheetId=file_id, range='Sheet1!A2:C4', valueInputOption='RAW', body=body
)
response = request.execute()

"""