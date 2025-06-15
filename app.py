from flask import Flask, request, jsonify, send_file
from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# === CONFIGURATION ===
SPREADSHEET_ID = '1s_joDeUNfCjk_U2UFON5ocDMcP3KjW7Y3WH7F1rmkd0'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'service_account.json'
EXTRACT_FOLDER = 'extracted'
os.makedirs(EXTRACT_FOLDER, exist_ok=True)

# === AUTHENTICATION ===
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
sheets_service = build('sheets', 'v4', credentials=credentials)

# === HELPERS ===
def get_sheet_names():
    meta = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    return [s['properties']['title'] for s in meta['sheets']]

def get_sheet_data(sheet):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=sheet
    ).execute()
    values = result.get('values', [])
    df = pd.DataFrame(values[1:], columns=values[0]) if values else pd.DataFrame()
    return df

def update_sheet(sheet, new_data):
    body = {'values': [new_data.columns.tolist()] + new_data.values.tolist()}
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=sheet,
        valueInputOption='RAW',
        body=body
    ).execute()

# === ROUTES ===
@app.route('/sheets', methods=['GET'])
def list_sheets():
    try:
        sheets = get_sheet_names()
        return jsonify(sheets)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/extract', methods=['POST'])
def extract():
    try:
        data = request.get_json()
        sheet_name = data['sheet']
        num_rows = int(data['count'])

        df = get_sheet_data(sheet_name)
        if df.empty:
            return jsonify({'error': 'Sheet is empty or not found.'}), 400
        if num_rows > len(df):
            return jsonify({'error': 'Not enough rows to extract.'}), 400

        extracted_df = df.head(num_rows).copy()
        remaining_df = df.iloc[num_rows:].copy()
        extracted_df.insert(0, 'SerialNumber', range(1, num_rows + 1))

        # Write to Excel
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{sheet_name}_{num_rows}_leads_{timestamp}.xlsx"
        filepath = os.path.join(EXTRACT_FOLDER, filename)

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            extracted_df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            for col_idx, column in enumerate(extracted_df.columns, start=1):
                col_letter = get_column_letter(col_idx)
                worksheet.column_dimensions[col_letter].width = 21.11 if col_idx == 2 else 15
                for row_idx in range(2, num_rows + 2):
                    cell = worksheet[f'{col_letter}{row_idx}']
                    if col_idx in [3, 7]:
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                    else:
                        cell.alignment = Alignment(horizontal='center', vertical='center')

        # Write remaining data back to Google Sheets
        update_sheet(sheet_name, remaining_df)

        return jsonify({'download_url': f'/download/{filename}'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    filepath = os.path.join(EXTRACT_FOLDER, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return jsonify({'error': 'File not found'}), 404

if __name__ == '__main__':
    app.run(debug=True)
