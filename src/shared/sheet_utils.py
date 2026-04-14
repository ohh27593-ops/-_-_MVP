from oauth2client.service_account import ServiceAccountCredentials
import gspread
import sys


def connect_google_sheet(credentials_file, sheet_id):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            credentials_file, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(sheet_id)
        return sheet
    except Exception as e:
        print('구글 시트 연결에 실패하여 코드를 자동으로 중지합니다. 에러 정보:', e)
        sys.exit()


def get_worksheet(sheet, worksheet_name):
    try:
        ws = sheet.worksheet(worksheet_name)
        return ws
    except Exception as e:
        print('워크시트 불러오기에 실패하여 코드를 자동으로 중지합니다. 에러 정보:', e)
        sys.exit()
