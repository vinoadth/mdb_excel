import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
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
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    drive_service = build('drive', 'v3', credentials=creds)

    folder_id = '1tbQ0zo7rbU32slosli9UhHlS1_sPo8Pl'

    service = build('sheets', 'v4', credentials=creds)

    personal = [
        ['Name', 'Age'],
        ['Vinoth', '32']
    ]

    official = [
        ['ID', 'Email'],
        ['1000', 'vinoth@mail.ru']
    ]

    spreadsheet = {
        'properties': {
            'title': "People"
        },
        'sheets': [
            {
                'properties': {
                    'title': 'Personal'
                }
            },
            {
                'properties': {
                    'title': 'Official'
                }
            }
        ]
    }

    batch_update_values_request_body = {
        'value_input_option': 'RAW',
        'data': [
            {
                'range': 'Personal!A1',
                'values': personal
            },
            {
                'range': 'Official!A1',
                'values': official
            }
        ],
    }

    spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                        fields='spreadsheetId').execute()
    file_id = spreadsheet.get('spreadsheetId')

    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=file_id,
        body=batch_update_values_request_body).execute()

    # Retrieve the existing parents to remove
    file = drive_service.files().get(fileId=file_id,
                                    fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    # Move the file to the new folder
    file = drive_service.files().update(fileId=file_id,
                                        addParents=folder_id,
                                        removeParents=previous_parents,
                                        fields='id, parents').execute()

if __name__ == '__main__':
    main()
