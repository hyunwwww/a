import requests
import os
from dotenv import load_dotenv
import config

# 티스토리 블로그의 카테고리를 불러내는 코드입니다.

# 환경 변수 로드
load_dotenv()

# 티스토리 API 인증 정보
access_token = os.getenv('ACCESS_TOKEN')
blog_name = config.blog_name

# 카테고리 목록 가져오기
url = 'https://www.tistory.com/apis/category/list'
params = {
    'access_token': access_token,
    'blogName': blog_name,
    'output': 'json'
}

response = requests.get(url, params=params)

# API 응답 검증 및 카테고리 목록 출력
if response.status_code == 200:
    categories = response.json().get('tistory', {}).get('item', {}).get('categories', [])
    for category in categories:
        print(f'Name: {category["name"]}, ID: {category["id"]}')
else:
    print("API 요청 실패:", response.status_code)