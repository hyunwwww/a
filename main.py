from googleapiclient.discovery import build
import gspread, os
from google.oauth2 import service_account
from tistory import Tistory
from dotenv import load_dotenv
import os
import config

# .env 파일에서 환경변수 불러오기
load_dotenv()
client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
access_token = os.getenv('ACCESS_TOKEN')
credentials = os.getenv('creds')        # 구글인증 환경변수(JSON)

# config.py에서 설정값 불러오기
blog_url = config.blog_url
SPREADSHEET_ID = config.SPREADSHEET_ID
RANGE_NAME = config.RANGE_NAME 

# 구글시트 API(JSON) 호출
SERVICE_ACCOUNT_FILE = 'C:/Users/LG/Documents/json_keys/peak-vista-382911-d8c6841af269.json'  # 서비스 계정 파일의 경로
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']   # 스프레드시트 읽기 및 쓰기 권한을 위한 스코프
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)                # 서비스 계정 인증 정보 생성
service = build('sheets', 'v4', credentials=creds)          # 구글 시트 서비스 빌드
sheet = service.spreadsheets()

# 티스토리 인스턴스 생성
ts = Tistory(blog_url, client_id, client_secret)
ts.access_token = access_token

# 구글 시트에서 데이터 가져오기
result = sheet.values().get(
    spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME
).execute()
values = result.get('values', [])

# 각 행에 대해 블로그 게시
for i, row in enumerate(values, start=1):
    # J열이 빈 셀인 경우를 고려하여 길이 확인
    if len(row) < 9 or row[8] != "완료":  # J열 확인
        title = row[0]     # B열: 제목

        # 내용 구성: G, D, H, E, F 열의 데이터
        content = f'<img src="{row[6]}">\n\n\n'  # G열: 메인 이미지 URL
        content += f'{row[3]}\n\n\n'             # D열: 게시물 본문
        content += f'<img src="{row[7]}">\n\n\n' # H열: 서브 이미지 URL
        content += f'{row[4]}\n\n\n'             # E열: 게시물 요약
        content += f'{row[5]}\n\n\n'             # F열: 태그

        # 블로그 게시물 작성 및 업로드
        res = ts.write_post(
                    title=title, 
                    content=content, 
                    visibility="3",     # 발행상태 (0: 비공개 - 기본값, 1: 보호, 3: 발행)
                    acceptComment="1"   # 댓글 허용 (0 , 1 - 기본값)
                )
        print(res)

        # 업로드 완료 표시
        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID, 
            range=f"J{i+1}", 
            valueInputOption="RAW", 
            body={"values": [["완료"]]}
        ).execute()
