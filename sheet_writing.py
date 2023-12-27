import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os
import openai
import config
import prompts
import time
from dotenv import load_dotenv


# Load environment variables from .env file
load_dotenv()

# Google Sheets API 설정
SERVICE_ACCOUNT_FILE = os.getenv('SERVICE_ACCOUNT_FILE')

# 필요한 스코프 추가
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',  # 스프레드시트 읽기 및 쓰기 권한
    'https://www.googleapis.com/auth/drive'          # Google Drive API 접근 권한
]

# 서비스 계정 인증 정보 생성
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# 구글 시트 서비스 빌드
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()

# gspread 클라이언트 초기화
client = gspread.authorize(creds)

# OpenAI API 설정
openai.api_key = os.getenv('OPENAI_API_KEY')

# Google Sheets에서 데이터 읽기
SPREADSHEET_ID = config.SPREADSHEET_ID
sheet = client.open_by_key(SPREADSHEET_ID).sheet1  # 스프레드시트 ID를 사용하여 열기
main_concept = sheet.acell('B2').value
sub_concept = sheet.acell('D2').value
rangename = config.RANGE_NAME.split('!')[1]  # 'Sheet1!B4:P200'에서 'B4:P200'만 추출
cells = sheet.range(rangename)

### 제목 생성 (B열) ###

# 마지막으로 데이터가 있는 행 찾기
def find_last_row(sheet):
    str_list = filter(None, sheet.col_values(2))  # 2번째 열(B열)의 값들을 가져옴
    return len(list(str_list)) + 1  # 마지막 행 번호 반환

row = find_last_row(sheet)  # 시작 행을 마지막 행 다음으로 설정
titles_to_create = []  # 생성할 제목을 저장할 리스트

# 제목 목록 생성 (한 번만 수행)
if not titles_to_create:
    title_prompt = prompts.TITLE_PROMPT_TEMPLATE.format(main_concept=main_concept, sub_concept=sub_concept)
    title_response = openai.ChatCompletion.create(
        model='gpt-3.5-turbo',
        messages=[{"role": "system", "content": title_prompt}]
    )
    titles = title_response.choices[0].message['content'].strip().split('\n')
    titles_to_create.extend([title.strip().lstrip('0123456789.-*# ') for title in titles if title.strip()])

    list_length = len(titles_to_create)
    print(titles_to_create)
    print(f"리스트에 생성된 제목의 수: {list_length}")


# 각 행에 순차적으로 채우기
for i in range(row, row + len(titles_to_create)):
    title_cell = f"B{i}"
    title = sheet.acell(title_cell).value
    if not title and i-row < len(titles_to_create):
        sheet.update_acell(title_cell, titles_to_create[i-row])
        time.sleep(2)
                
    ### 본문 글 작성 (D열) ###
    title_cell = f"B{i}"
    title = sheet.acell(title_cell).value
    body_cell = f"D{i}"
    existing_data = sheet.acell(body_cell).value

    if title and not existing_data:

        # 본문 프롬프트 생성 및 API 호출
        body_prompt = prompts.BODY_PROMPT_TEMPLATE.format(
            title=title, main_concept=main_concept
        )
        body_response = openai.ChatCompletion.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": body_prompt}]
        )
        body = body_response.choices[0].message['content'].strip()
        
        # Google Sheets에 본문 기록
        sheet.update_acell(body_cell, body) # 열 D에 데이터를 기록합니다
        time.sleep(2)

    ### 요약 작성 시작 (E열) ###

    # 열 D에 내용 가져오기
    body_cell = f"D{i}"
    body = sheet.acell(body_cell).value

    # E열 기존 데이터 확인
    summary_cell = f"E{i}"
    existing_data = sheet.acell(summary_cell).value

    if not existing_data and body: 
        # OpenAI API를 사용하여 열 D에 들어갈 내용 생성
        summary_prompt = prompts.SUMMARY_PROMPT_TEMPLATE.format(content=body)
        summary_response = openai.ChatCompletion.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": summary_prompt}]
        )
        summary = summary_response.choices[0].message['content'].strip()

        # 열 E에 내용 기록
        sheet.update_acell(summary_cell, summary)
        time.sleep(2)


    ### 태그 작성 (F열) ###

    # 열 D에서 내용 가져오기
    body_cell = f"D{i}"
    body = sheet.acell(body_cell).value

    # F열 기존 데이터 확인
    tag_cell = f"F{i}"
    existing_data = sheet.acell(tag_cell).value

    if not existing_data and body: 
        # OpenAI API를 사용하여 열 F에 들어갈 내용 생성
        tag_prompt = prompts.TAG_PROMPT_TEMPLATE.format(content=body)
        tag_response = openai.ChatCompletion.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": tag_prompt}]
        )
        tag = tag_response.choices[0].message['content'].strip()

        # 열 F에 내용 기록
        sheet.update_acell(tag_cell, tag)
        time.sleep(2)

    ### 카테고리 (C열) ###

    # 열 D에서 내용 가져오기
    body_cell = f"D{i}"
    body = sheet.acell(body_cell).value

    # F열 기존 데이터 확인
    category_cell = f"C{i}"
    existing_data = sheet.acell(category_cell).value

    if not existing_data and body: 
        # OpenAI API를 사용하여 열 C에 들어갈 내용 생성
        category_prompt = prompts.CATEGORY_PROMPT_TEMPLATE.format(content=body)
        category_response = openai.ChatCompletion.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": category_prompt}]
        )
        category = category_response.choices[0].message['content'].strip()

        # 열 C에 내용 기록
        sheet.update_acell(category_cell, category)
        time.sleep(2)

    ### 인트로멘트 (J열) ###

    # J열 기존 데이터 확인
    ment_intro_cell = f"J{i}"
    existing_data = sheet.acell(ment_intro_cell).value

    if not existing_data : 
        # OpenAI API를 사용하여 열 J에 들어갈 내용 생성
        ment_intro_prompt = prompts.MENT_INTRO.format(content=body)  # 프롬프트 내 {body} 의 내용참조
        ment_intro_response = openai.ChatCompletion.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": ment_intro_prompt}]
        )
        ment_intro = ment_intro_response.choices[0].message['content']

        # 열 J에 내용 기록
        sheet.update_acell(ment_intro_cell, ment_intro)
        time.sleep(2)

    ### 카테고리 (K열) ###

    # F열 기존 데이터 확인
    closing_cell = f"K{i}"
    existing_data = sheet.acell(closing_cell).value

    if not existing_data : 
        # OpenAI API를 사용하여 열 K에 들어갈 내용 생성
        closing_prompt = prompts.MENT_CLOSING.format(content=body)  # 프롬프트 내 {body} 의 내용참조 
        closing_response = openai.ChatCompletion.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": closing_prompt}]
        )
        closing = closing_response.choices[0].message['content']

        # 열 K에 내용 기록
        sheet.update_acell(closing_cell, closing)
        time.sleep(2)


        
 


