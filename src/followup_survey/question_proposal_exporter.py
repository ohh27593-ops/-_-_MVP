import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import sys


# 1. 저장 경로 설정
way_question = r"data/user_questions.xlsx"
way_proposal = r"data/user_proposals.xlsx"
try:
    compare_data_0 = pd.read_excel(r"data/user_info.xlsx")
except Exception as ereo:
    print('이전 사용자 정보 가져오기에 실패하여 코드를 자동으로 중지합니다. 에러 정보:', ereo)
    sys.exit()

# 2. 인증 및 시트 연결
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    r"credentials.json", scope)
client = gspread.authorize(creds)

# 3. 시트 가져오기
sheet = client.open_by_key("your_second_survey_sheet_id")

# 4. 전체 워크시트 이름 확인
worksheets = sheet.worksheets()
print('4. 전체 워크시트 이름 확인', 'worksheets:', worksheets)

# 5. 특정 워크시트 선택
ws = sheet.worksheet("설문지 응답 시트2")
all_records = ws.get_all_records()
google_uni = pd.DataFrame(all_records)

# 타임스탬프 pd.Datetime으로 변환
def day_time_sleep(day_time_str):
    if "오전" in day_time_str:
        day_time_str = day_time_str.replace("오전", "AM")
    elif "오후" in day_time_str:
        day_time_str = day_time_str.replace("오후", "PM")
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

try:
    print(google_uni)
    google_uni = google_uni[
        (google_uni["친구 매칭 참여 여부"] == "네, 참여합니다.") &
        (google_uni['휴대전화 번호를 입력해주세요'].isin(
            compare_data_0.loc[compare_data_0['integ'].fillna(False) == True, 'phone']
        ))
    ]
    print('google_uni파일 기존 사용자 정보와 대조 완료')
    print(compare_data_0)
    print(google_uni)
except Exception as e:
    print(f'google_uni파일 기존 사용자 정보와 대조 중 오류: {e}')
    sys.exit()

to_go_excell_3 = {}
to_go_excell_3['stamp'] = google_uni['타임스탬프'].apply(day_time_sleep)
to_go_excell_3['any_more'] = google_uni['추가 활동 제안 ']
to_go_excell_3 = pd.DataFrame(to_go_excell_3)

try:
    to_go_excell_3 = to_go_excell_3[to_go_excell_3['any_more'] != '']
    to_go_excell_3.to_excel(way_proposal, index=False)
except Exception as e:
    print(f'to_go_excell의 query에서 오류 {e}')
    sys.exit()

to_go_excell_2 = {}
to_go_excell_2['stamp'] = google_uni['타임스탬프'].apply(day_time_sleep)
to_go_excell_2['phone'] = google_uni['휴대전화 번호를 입력해주세요']
to_go_excell_2['question'] = google_uni['친구 매칭에 관해 궁금한 점 ']
to_go_excell_2 = pd.DataFrame(to_go_excell_2)

try:
    to_go_excell_2 = to_go_excell_2[to_go_excell_2['question'] != '']
    to_go_excell_2.to_excel(way_question, index=False)
except Exception as e:
    print(f'to_go_excell의 query에서 오류 {e}')
    sys.exit()

print('to_go_excell_2')
print(to_go_excell_2)
print('to_go_excell_3')
print(to_go_excell_3)
print('작업 완료')
