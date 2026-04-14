import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import sys
import numpy as np


# 1. 비교용 지난 자료 가져오기
way_excel = r"data/user_info.xlsx"
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
sheet = client.open_by_key("your_first_survey_sheet_id")

# 4. 전체 워크시트 이름 확인
worksheets = sheet.worksheets()
print('4. 전체 워크시트 이름 확인', 'worksheets:', worksheets)

# 5. 특정 워크시트 선택 (첫 번째)
ws = sheet.worksheet("설문지 응답 시트1")
google_uni = pd.DataFrame(ws.get_all_records())

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
    .drop(columns=['타임스탬프', '성함을 입력해주세요 (실명 기준) ',
                                          '휴대전화 번호를 입력해주세요 ',
                                          '복돌복실의 운영 원칙에 동의하시나요?  ',
                                          '개인정보 수집 및 이용  '])\
    .sort_values('stamp', ascending=False)
# stamp: submission_timestamp 입력시점(사용자가 pdps에 구글 폼으로 가입을 신청한 시점)
# phone: mobile_phone_number(사용자의 휴대전화 번호)
# name: user_name: 사용자의 이름
# integ: integrity: 건전성(pdps사용자의 행위가 운영 정책에 위배되는가. 위배되지 않음:True, 위배됨:False)

# 7. compare_data와의 비교, 자료 에러 확인
compare = compare_data.loc[0, 'stamp']
recent_submission_data = to_book_data.query('stamp>@compare')
compare_data_2 = to_book_data.query('stamp<=@compare').sort_values('stamp', ascending=False)
try:
    compagr_data_merged = pd.merge(compare_data, compare_data_2, how='left', on=['stamp', 'phone', 'name'])
    print('7. compare_data와의 비교, 자료 에러 확인 결과 정상(merge가 정상적으로 됨-> stamp,phone,name 값 동일)')
except Exception as e:
    print('7. compare_data와의 비교, 자료 에러 확인 도중 merge불가. 에러발생:', e)
    sys.exit()

# 8. pdps 기존 가입자정보 확인(중복가입 방지)
if len(recent_submission_data) >= 1:
    try:
        recent_submission_data = recent_submission_data.assign(
            dupli=recent_submission_data['phone'].isin(compare_data['phone'])
        )
        recent_submission_data = recent_submission_data.query('dupli==False')
    except Exception as e:
        print('8. pdps 기존 가입자정보 확인(중복가입 방지) 중 에러발생:', e)
        print(recent_submission_data)
        sys.exit()
# dupli: duplicate sign-up(중복가입) (중복가입 맞음:True, 중복가입 아님:False)

# 9. pdps 탈퇴자 정보 삭제(불건전 이용자의 정보 삭제 대상에서 제외)
way_uninteg_user = r"data/uninteg_user.xlsx"
try:
    uninteg_user = pd.read_excel(way_uninteg_user)
except Exception as ero_1:
    print('uninteg_user 사용자 정보 가져오기에 실패하여 코드를 자동으로 중지합니다. 에러 정보:', ero_1)
    sys.exit()
compare_data['integ'] = np.where(compare_data['phone'].isin(uninteg_user['phone']), False, True)
# uninteg_user 파일을 기반으로 새로 uninteg에 등록된 사용자의 integ열의 정보를 False로 바꾼다.

way_delete_user = r"data/delete_user.xlsx"
try:
    delete_user = pd.read_excel(way_delete_user)
except Exception as ero_2:
    print('way_delete_user 사용자 정보 가져오기에 실패하여 코드를 자동으로 중지합니다. 에러 정보:', ero_2)
    sys.exit()

if len(recent_submission_data) >= 1:
    mask = recent_submission_data['phone'].isin(delete_user['phone'])
    recent_submission_data = recent_submission_data.loc[~mask].copy()

# 10. 신규 가입자 저장
new_user_way = r"data/new_user.xlsx"
recent_submission_data.to_excel(new_user_way, index=False)

print('recent_submission_data')
print(recent_submission_data)
print('compagr_data_merged')
print(compagr_data_merged)
print('compare_data')
print(compare_data)
print('작업 완료')
print(google_uni)
