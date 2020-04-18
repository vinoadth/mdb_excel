import os, sys, subprocess, time
from flask import Flask, flash, request, redirect, render_template
import xlsxwriter
from werkzeug.utils import secure_filename
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/uploads'

ALLOWED_EXTENSIONS = {'mdb'}

SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

google_sheet_url = 'https://docs.google.com/spreadsheets/d/'

creds = None

if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # if user does not select file, browser also
        # submit an empty part without filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            DATABASE = app.config['UPLOAD_FOLDER'] + '/' + filename
            while not os.path.exists(DATABASE):
                time.sleep(1)
            table_names = subprocess.Popen(['mdb-tables', '-1', DATABASE], stdout=subprocess.PIPE).communicate()[0]
            if isinstance(table_names, bytes):
                table_names = table_names.decode("utf-8")
            tables = table_names.split('\n')

            drive_service = build('drive', 'v3', credentials=creds)

            folder_id = '1tbQ0zo7rbU32slosli9UhHlS1_sPo8Pl'

            service = build('sheets', 'v4', credentials=creds)

            spreadsheet = {
                'properties': {
                    'title': filename
                },
                'sheets': []
            }

            batch_update_values_request_body = {
                'value_input_option': 'RAW',
                'data': []
            }

            for table in tables:
                if table != '':
                    obj = {}
                    obj['properties'] = {'title': table}
                    spreadsheet['sheets'].append(obj)

            for table in tables:
                if table != '':
                    # filename = table.replace(' ','_') + '.csv'
                    output = subprocess.Popen(['mdb-export', DATABASE, table], stdout=subprocess.PIPE).communicate()
                    raw_text = output[0].decode("utf-8")
                    rows = raw_text.split('\n')
                    excel_rows = []
                    for row in rows:
                        cells = row.split(',')
                        if len(cells) == 1 and cells[0] == '':
                            continue
                        for ind, cell in enumerate(cells):
                            if '"' in cell:
                                cells[ind] = cell.replace('"', '')
                                cells[ind] = cells[ind].strip()
                        excel_rows.append(cells)
                    result = {'range': (table + '!A1'), 'values': excel_rows}
                    batch_update_values_request_body['data'].append(result)

            spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                                fields='spreadsheetId').execute()
            file_id = spreadsheet.get('spreadsheetId')

            result = service.spreadsheets().values().batchUpdate(
                spreadsheetId=file_id,
                body=batch_update_values_request_body).execute()

            file = drive_service.files().get(fileId=file_id,
                                            fields='parents').execute()
            previous_parents = ",".join(file.get('parents'))

            file = drive_service.files().update(fileId=file_id,
                                                addParents=folder_id,
                                                removeParents=previous_parents,
                                                fields='id, parents').execute()

            os.remove(DATABASE)
            result_url = google_sheet_url + file_id
            return redirect(result_url)

    return render_template('index.html')
