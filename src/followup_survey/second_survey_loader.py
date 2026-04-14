import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import sys
import numpy as np


# 1. 비교용 지난 자료 가져오기
way_excel = r"data/user_info_second_survey.xlsx"
way_question = r"data/user_questions.xlsx"
way_proposal = r"data/user_proposals.xlsx"
try:
    compare_data = pd.read_excel(way_excel)
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

# 5. 특정 워크시트 선택 (첫 번째)
ws = sheet.worksheet("설문지 응답 시트2")
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
to_go_excell = {}
to_go_excell['stamp'] = to_go_excell_3['stamp']
to_go_excell_3['any_more'] = google_uni['추가 활동 제안 ']
to_go_excell_3 = pd.DataFrame(to_go_excell_3)
try:
    to_go_excell_3 = to_go_excell_3[to_go_excell_3['any_more'] != '']
    to_go_excell_3.to_excel(way_proposal, index=False)
except Exception as e:
    print(f'to_go_excell의 query에서 오류 {e}')
    sys.exit()


to_go_excell['phone'] = google_uni['휴대전화 번호를 입력해주세요']
to_go_excell['monday'] = google_uni['참여 가능한 시간대  [월요일]']
to_go_excell['tuesday'] = google_uni['참여 가능한 시간대  [화요일]']
to_go_excell['wednesday'] = google_uni['참여 가능한 시간대  [수요일]']
to_go_excell['thursday'] = google_uni['참여 가능한 시간대  [목요일]']
to_go_excell['friday'] = google_uni['참여 가능한 시간대  [금요일]']
to_go_excell['saturday'] = google_uni['참여 가능한 시간대  [토요일]']
to_go_excell['sunday'] = google_uni['참여 가능한 시간대  [일요일]']
to_go_excell['prefer'] = google_uni['선호하는 활동 ']
to_go_excell = pd.DataFrame(to_go_excell)

to_go_excell_2 = {}
to_go_excell_2['stamp'] = to_go_excell['stamp']
to_go_excell_2['phone'] = to_go_excell['phone']
to_go_excell_2['question'] = google_uni['친구 매칭에 관해 궁금한 점 ']
to_go_excell_2 = pd.DataFrame(to_go_excell_2)

print('to_go_excell')
print(to_go_excell)
print('to_go_excell_3')
print(to_go_excell_3)

# 0) to_go_excell과 compare_data concat
frames = [df for df in (compare_data, to_go_excell) if not df.empty]
to_go_excell = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

# 1) 최신 stamp가 마지막이 되도록 정렬
to_go_excell = to_go_excell.sort_values(['phone', 'stamp'])

# 2) phone 기준으로 중복 제거: 마지막(최신)만 유지
to_go_excell = (
    to_go_excell
      .drop_duplicates(subset='phone', keep='last')
      .reset_index(drop=True)
)

try:
    to_go_excell_2 = to_go_excell_2[to_go_excell_2['question'] != '']
    to_go_excell_2.to_excel(way_question, index=False)
except Exception as e:
    print(f'to_go_excell의 query에서 오류 {e}')
    sys.exit()

to_go_excell.to_excel(way_excel, index=False)



days = ['monday','tuesday','wednesday','thursday','friday','saturday','sunday']

for col in days:
    be_match = to_go_excell[['phone', f'{col}', 'prefer']]

    be_match = be_match.assign(morning=np.where(be_match[f'{col}'].astype('string').str.contains('오전', regex=False, na=False), True, False),
                             lunch=np.where(be_match[f'{col}'].astype('string').str.contains('점심', regex=False, na=False), True, False),
                             afternoon=np.where(be_match[f'{col}'].astype('string').str.contains('오후', regex=False, na=False), True, False),
                             cafe=np.where(be_match['prefer'].astype('string').str.contains('카페', regex=False, na=False),
                                  True, False),
                             walk=np.where(be_match['prefer'].astype('string').str.contains('산책', regex=False, na=False),
                                  True, False),
                             restaurant=np.where(be_match['prefer'].astype('string').str.contains('맛집', regex=False, na=False),
                                  True, False),
                             recreation=np.where(be_match['prefer'].astype('string').
                                                 str.contains('놀이(보드게임 등)', regex=False, na=False),
                                  True, False),
                             exercise=np.where(be_match['prefer'].astype('string').str.contains('운동', regex=False, na=False),
                                  True, False)
                             )
    be_match = be_match.drop(columns=['prefer'])

    # 3) 필터링 후 요일별 파일로 저장
    be_match.to_excel(
        fr"data/user_info_2_{col}.xlsx",
        index=False
    )
