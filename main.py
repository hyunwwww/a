from googleapiclient.discovery import build
import gspread, os
from google.oauth2 import service_account
from tistory import Tistory
from dotenv import load_dotenv
import os
import config
import requests
import markdown

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
blogname = config.blog_name

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
    # L열(완료열)이 빈 셀인 경우를 고려하여 길이 확인
    if len(row) < 11 or row[10] != "완료":  # L열 확인
        
        title = row[0]     # B열: 제목

        # 전체 콘텐츠를 마크다운 형식으로 작성
        content = ""

        # 서두 (마크다운 형식)
        markdown_introduction = row[8]  # 서두
        content += markdown.markdown(markdown_introduction)
        content += '\n\n'
        content += '<br>' * 5

        # 메인 이미지 (HTML 형식)
        content += f'  <img src="{row[5]}" style="display: block; margin: auto; width: 90%;" />\n'  # 메인 이미지 URL
        content += '<br>' * 5

        # 본문 (마크다운 형식)
        markdown_body = row[2]  # 마크다운 본문
        content += markdown.markdown(markdown_body)
        content += '\n\n'
        content += '<br>' * 5

        # 부가 이미지 1 (HTML 형식)
        content += f'<img src="{row[6]}" style="display: block; margin: auto; width: 75%;" />\n'  # 부가 이미지 1 URL
        content += '<br>' * 5

        # 요약 (마크다운 형식)
        markdown_summary = row[3]  # 요약
        content += markdown.markdown(markdown_summary)
        content += '\n\n'
        content += '<br>' * 5

        # 부가 이미지 2 (HTML 형식)
        content += f'<img src="{row[7]}" style="display: block; margin: auto; width: 75%;" />\n'  # 부가 이미지 2 URL
        content += '<br>' * 5

        # 마무리 코멘트 (마크다운 형식)
        markdown_closing_comment = row[9]  # 마무리 코멘트
        content += markdown.markdown(markdown_closing_comment)
        content += '\n\n'
        content += '<br>' * 3
        
        tags = row[4] # tags            


        # 카테고리 관련 파라메터 요청
        # 카테고리 목록 가져오기
        url = 'https://www.tistory.com/apis/category/list'
        params = { 
            'access_token': os.getenv('ACCESS_TOKEN'),
            'output': 'json' ,
            'blogName': config.blog_name,

        }

        response = requests.get(url, params=params)

        # API 응답 검증
        if response.status_code == 200:
            categories = response.json().get('tistory', {}).get('item', {}).get('categories', [])
        else:
            print("API 요청 실패:", response.status_code)
            categories = []

        # 구글 시트에서 카테고리명 읽기
        # 예: sheet_category_name = '여행' (이 부분은 구글 시트 API를 사용하여 구현해야 함)
        sheet_category_name = row[1]  # 이 부분은 실제 구글 시트에서 읽어온 값으로 대체해야 함

        # 티스토리 카테고리 ID 찾기
        category_id = None
        for category in categories:
            if category['name'] == sheet_category_name:
                category_id = category['id']
                break

        if category_id:
            print(f'카테고리 ID: {category_id}')
        else:
            print('일치하는 카테고리가 없습니다.')

        # 블로그 게시물 작성 및 업로드
        res = ts.write_post(
                    title=title, 
                    content=content,                     
                    visibility="0",     # 발행상태 (0: 비공개 - 기본값, 1: 보호, 3: 발행)
                    acceptComment="1",   # 댓글 허용 (0 , 1 - 기본값)
                    category=category_id,
                    tag=tags
        )    

        print(res)

        # 업로드 완료 표시
        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID, 
            range=f"Sheet1!L{i+3}",  # 완료표기 시작되는 올바른 셀 위치 지정
            valueInputOption="RAW", 
            body={"values": [["완료"]]}
        ).execute()



