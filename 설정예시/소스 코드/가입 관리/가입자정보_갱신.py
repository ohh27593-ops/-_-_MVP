import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import sys
import numpy as np


# 1. 비교용 지난 자료 가져오기
way_excel = r"데이터/user_info.xlsx"
try:
    compare_data = pd.read_excel(way_excel)
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
sheet = client.open_by_key("여기에_1차설문_구글시트_ID_입력")

# 4. 전체 워크시트 이름 확인
worksheets = sheet.worksheets()
print('4. 전체 워크시트 이름 확인','worksheets:', worksheets)

# 5. 특정 워크시트 선택 (첫 번째)
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

# 6. google_sheet 데이터 전처리
google_uni['타임스탬프'] = google_uni['타임스탬프'].apply(day_time_sleep)
to_book_data = google_uni.assign(stamp=google_uni['타임스탬프'],
                  phone=google_uni['휴대전화 번호를 입력해주세요 '],
                  name=google_uni['성함을 입력해주세요 (실명 기준) '],
                  integ=True)\
    .drop(columns=['타임스탬프','성함을 입력해주세요 (실명 기준) ',
                                          '휴대전화 번호를 입력해주세요 ',
                                          '복돌복실의 운영 원칙에 동의하시나요?  ',
                                          '개인정보 수집 및 이용  '])\
    .sort_values('stamp', ascending=False)
# stamp: submission_timestamp 입력시점(사용자가 pdps에 구글 폼으로 가입을 신청한 시점)
# phone: mobile_phone_number(사용자의 휴대전화 번호)
# name: user_name: 사용자의 이름
# integ: integrity: 건전성(pdps사용자의 행위가 운영 정책에 위배되는가. 위배되지 않음:True, 위배됨:False)

# 8-0. 당일 중복가입 방지 확인
to_book_data = to_book_data.sort_values(['phone', 'stamp'])

to_book_data = (
    to_book_data
      .drop_duplicates(subset='phone', keep='last')
      .reset_index(drop=True)
)
print('to_book_data')
print(to_book_data)

# 8. pdps 기존 가입자정보 확인(중복가입 방지) 후 가입 완료
to_book_data = to_book_data.drop(
    to_book_data[to_book_data['phone'].isin(compare_data['phone'])].index
)
compare_data = pd.concat([compare_data, to_book_data], ignore_index=True)   # 가입 완료

# 9. pdps 탈퇴자 정보 삭제(불건전 이용자의 경우 정보 삭제 대상에서 제외)
way_uninteg_user = r"데이터/uninteg_user.xlsx"
try:
    uninteg_user = pd.read_excel(way_uninteg_user)
except Exception as ero_1:
    print('uninteg_user 사용자 정보 가져오기에 실패하여 코드를 자동으로 중지합니다. 에러 정보:', ero_1)
    sys.exit()
compare_data['integ'] = np.where(compare_data['phone'].isin(uninteg_user['phone']), False, True)
# uninteg_user 파일을 기반으로 새로 uninteg에 등록된 사용자의 integ열의 정보를 False로 바꾼다.

way_delete_user = r"데이터/delete_user.xlsx"
try:
    delete_user = pd.read_excel(way_delete_user)
except Exception as ero_2:
    print('way_delete_user 사용자 정보 가져오기에 실패하여 코드를 자동으로 중지합니다. 에러 정보:', ero_2)
    sys.exit()

mask = compare_data['phone'].isin(delete_user['phone']) & (compare_data['integ'] == True)
compare_data = compare_data.loc[~mask].copy()
# 또는
# compare_data = compare_data.drop(compare_data[mask].index)

# 10. 새로운 compare_data excel파일에 저장
compare_data.to_excel(way_excel, index=False)

# 11. delete_user정보 삭제
# 모든 행 삭제, 컬럼(헤더)은 유지
delete_user = delete_user.iloc[0:0]
delete_user.to_excel(way_delete_user, index=False)

# 11. recent_submission_data(새로운 가입자) 카카오톡에 인증톡 보내기 (이건 내가 수동으로..ㅎㅎ)
# 카카오톡으로 보내야 하는 list만 excell에 나타내 주기
print('새로 가입한 가입자 정보')
print(to_book_data)
new_user_way = r"데이터/new_user.xlsx"
to_book_data.to_excel(new_user_way, index=False)
# 반드시 이들에게 수동으로 카카오톡 안내 메시지를 보내야 한다!!
