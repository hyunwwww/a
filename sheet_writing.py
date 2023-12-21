import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os
import openai
import config
import prompts




# Google Sheets API 설정
SERVICE_ACCOUNT_FILE = 'C:/Users/LG/Documents/json_keys/peak-vista-382911-d8c6841af269.json'  # 서비스 계정 파일의 경로

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


# 새로운 포스트 작성 및 업데이트
for i in range(0, len(cells), 11):  # 11개의 열로 구성된 행을 순회
    if cells[i].value == '':  # 첫 번째 열 (B열)이 비어있는 경우
        
        # 제목 프롬프트 생성 및 API 호출
        title_prompt = prompts.TITLE_PROMPT_TEMPLATE.format(
            main_concept=main_concept, 
            sub_concept=sub_concept
        )
        title_response = openai.Completion.create(
            engine="text-davinci-003",
            prompt=title_prompt, 
            max_tokens=140
        )
        title = title_response.choices[0].text.strip()

        # 제목을 B열에 먼저 기록
        cells[i].value = title

        # 본문 프롬프트 생성 및 API 호출
        body_prompt = prompts.BODY_PROMPT_TEMPLATE.format(
            title=title, 
            main_concept=main_concept, 
            sub_concept=sub_concept
        )
        body_response = openai.Completion.create(
            engine="text-davinci-003",              
            prompt=body_prompt, 
            max_tokens=3000
        )
        body = body_response.choices[0].text.strip()

        # 본문을 D열에 기록
        cells[i+2].value = body  # D열에 본문 기록

        # 제목과 본문이 모두 기록된 후에 변경 사항 저장
        sheet.update_cells(cells)
        break  # 하나의 포스트만 작성
