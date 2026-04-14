import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import sys


# 1. 인증 및 시트 연결
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    r"credentials.json", scope)
client = gspread.authorize(creds)

# 2. 시트 가져오기
sheet = client.open_by_key("your_first_survey_sheet_id")

# 3. 전체 워크시트 이름 확인
worksheets = sheet.worksheets()
print('3. 전체 워크시트 이름 확인', 'worksheets:', worksheets)

# 4. 특정 워크시트 선택 (첫 번째)
ws = sheet.worksheet("설문지 응답 시트1")
all_records = ws.get_all_records()
google_uni = pd.DataFrame(all_records)

# 타임스탬프 pd.Datetime으로 변환
def day_time_sleep(day_time_str):
    # 1. 한글 오전/오후 → 영어 AM/PM으로 바꾸기
    if "오전" in day_time_str:
        day_time_str = day_time_str.replace("오전", "AM")
    elif "오후" in day_time_str:
        day_time_str = day_time_str.replace("오후", "PM")
    # 2. datetime 변환
    try:
        dt = pd.to_datetime(day_time_str, format='%Y. %m. %d %p %I:%M:%S')
        return dt
    except Exception as ero:
        print("타임스탬프 pd.Datetime으로 변환 실패 코드 중지. 에러:", ero)
        sys.exit()

def call_phone_number(call_phone):
    call_num = str(call_phone)
    call_num = '0' + call_num
    return call_num

# 5. google_sheet 데이터 전처리
google_uni['타임스탬프'] = google_uni['타임스탬프'].apply(day_time_sleep)
to_book_data = google_uni.assign(
    stamp=google_uni['타임스탬프'],
    phone=google_uni['휴대전화 번호를 입력해주세요 '],
    name=google_uni['성함을 입력해주세요 (실명 기준) '],
    integ=True
).drop(columns=[
    '타임스탬프',
    '성함을 입력해주세요 (실명 기준) ',
    '휴대전화 번호를 입력해주세요 ',
    '복돌복실의 운영 원칙에 동의하시나요?  ',
    '개인정보 수집 및 이용  '
]).sort_values('stamp', ascending=False)

# stamp: submission_timestamp 입력시점(사용자가 pdps에 구글 폼으로 가입을 신청한 시점)
# phone: mobile_phone_number(사용자의 휴대전화 번호)
# name: user_name: 사용자의 이름
# integ: integrity: 건전성(pdps사용자의 행위가 운영 정책에 위배되는가. 위배되지 않음:True, 위배됨:False)

print('google_uni')
print(google_uni)
print('to_book_data')
print(to_book_data)
print('작업 완료')
