import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os
from openai import OpenAI
import config
import prompts
import time
from dotenv import load_dotenv
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import subprocess
from selenium.webdriver.chrome.options import Options

#### 전체 설정 부분 ####

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
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

# 구글 시트 서비스 빌드
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()

# gspread 클라이언트 초기화
gspread_client = gspread.authorize(creds)

# OpenAI API 설정
openai_client = OpenAI(api_key = os.getenv('OPENAI_API_KEY'))

# Google Sheets에서 데이터 읽기
SPREADSHEET_ID = config.SPREADSHEET_ID
sheet = gspread_client.open_by_key(SPREADSHEET_ID).sheet1  # 스프레드시트 ID를 사용하여 열기
main_concept = sheet.acell('B2').value
sub_concept = sheet.acell('D2').value
rangename = config.RANGE_NAME.split('!')[1]  # 'Sheet1!B4:P200'에서 'B4:P200'만 추출
cells = sheet.range(rangename)

openai_client = OpenAI()


def duckduckgo_image_search(query):
    driver = None
    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

        chrome_driver_path = r"C:\chrome-120.0.6099.109\120\chromedriver.exe"
        service = Service(executable_path=chrome_driver_path)

        subprocess.Popen(
            r'C:\chrome-120.0.6099.109\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrometemp"'
        )
        print("Chrome 디버깅 모드 실행됨")

        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("WebDriver 인스턴스 생성됨")
        time.sleep(3)

        image_info = []

        for i in range(10):
            search_url = f"https://duckduckgo.com/?q={query}&ia=images&iax=images"
            driver.get(search_url)

            all_image_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".tile--img__img.js-lazyload"))
            )
            image_elements = all_image_elements[i]

            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(image_elements))
            driver.execute_script("arguments[0].click();", image_elements)
            time.sleep(2)  # 클릭 후에 잠시 기다림

            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')

            link = soup.find('a', {'class': 'detail__media__img-link.js-detail-img'})
            image_src = link.get('href') if link else None
            if not image_src:
                img_tag = soup.find('img', {'class': 'detail__media__img-highres.js-detail-img.js-detail-img-high'})
                if img_tag:
                    image_src = img_tag.get('src')
                    # 'https:'가 없는 경우 추가
                    if image_src and not image_src.startswith('https:'):
                        image_src = 'https:' + image_src
                        
            title_tag = soup.find('span', {'class': 'tile--img__title'})
            image_title = title_tag.text if title_tag else None
            image_size_tag = soup.find('div', {'class': 'tile--img__dimensions'})
            image_size = image_size_tag.text if image_size_tag else None

            image_info.append({
                'url': image_src,
                'title': image_title,
                'size': image_size
            })

        driver.back()
                        
        print(image_info)

        return image_info

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        if driver:
            driver.quit()
            print("드라이버 세션 종료됨")



# # 이미지 첫 번째 프롬프트 : 웹에 검색할 단어 선택(검색하는 데 까지 관여)
# # 현재미사용
# def select_relevant_images(images, content):
#     prompt = prompts.IMAGE_TEMPLATE_1.format(content=content)
#     for i, img in enumerate(images):
#         prompt += f"{i+1}: URL: {img}\n"

#     response = openai_client.chat.completions.create(
#         model="gpt-3.5-turbo",
#         messages=[{"role": "system", "content": prompt}],
#         max_tokens=150  # 토큰 수를 증가시켜 여러 이미지를 선택할 수 있도록 함
#     )

#     # ChatGPT의 응답에서 세 개의 이미지 URL 추출
#     response_text = response.choices[0].message.content.strip().split('\n')
#     selected_urls = [url for url in response_text if url.startswith('http')]
#     return selected_urls[:3]  # 상위 세 개의 URL 반환


# 이미지2번째 프롬프트 : 이미지 정보를 OpenAI API를 사용하여 본문글과 비교,
# 세개를 선택하여 목록을 제출
def analyze_images_with_openai(image_info, body_content):
    image2_prompt = prompts.IMAGE_TEMPLATE_2.format(
        body_content=body_content, 
        image_info = image_info
    )

    image2_response = openai_client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "system", "content": image2_prompt}],
    )

    response_text = image2_response.choices[0].message.content.strip('')
    urls = response_text.split('\n')  # URL은 각 줄에 하나씩 있다고 가정

    return urls

# 마지막 데이터가 있는 행 찾기
def find_last_row(sheet):
    str_list = filter(None, sheet.col_values(2))  # 2번째 열(B열)의 값들을 가져옴
    return len(list(str_list)) + 1  # 마지막 행 번호 반환







### 제목 생성 (B열) ###

row = find_last_row(sheet)  # 시작 행을 마지막 행 다음으로 설정
titles_to_create = []  # 생성할 제목을 저장할 리스트

# 제목 목록 생성 (한 번만 수행)
if not titles_to_create:
    title_prompt = prompts.TITLE_PROMPT_TEMPLATE.format(main_concept=main_concept, sub_concept=sub_concept)
    title_response = openai_client.chat.completions.create(
        model='gpt-3.5-turbo',
        messages=[{"role": "system", "content": title_prompt}]
    )
    titles = title_response.choices[0].message.content.strip().split('\n')
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
        time.sleep(1)

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
        body_response = openai_client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": body_prompt}]
        )
        body = body_response.choices[0].message.content.strip()

        # Google Sheets에 본문 기록
        sheet.update_acell(body_cell, body) # 열 D에 데이터를 기록합니다
        time.sleep(1)

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
        summary_response = openai_client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": summary_prompt}]
        )
        summary = summary_response.choices[0].message.content.strip()

        # 열 E에 내용 기록
        sheet.update_acell(summary_cell, summary)
        time.sleep(1)


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
        tag_response = openai_client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": tag_prompt}]
        )
        tag = tag_response.choices[0].message.content.strip()

        # 열 F에 내용 기록
        sheet.update_acell(tag_cell, tag)
        time.sleep(1)

    ### 카테고리 (C열) ###

    # 열 D에서 내용 가져오기
    body_cell = f"D{i}"
    body = sheet.acell(body_cell).value

    # F열 기존 데이터 확인
    category_cell = f"C{i}"
    existing_data = sheet.acell(category_cell).value

    if not existing_data and body: 
        # OpenAI API를 사용하여 열 C에 들어갈 내용 생성
        category_prompt = prompts.CATEGORY_PROMPT_TEMPLATE.format(content = body)
        category_response = openai_client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": category_prompt}]
        )
        category = category_response.choices[0].message.content.strip()

        # 열 C에 내용 기록
        sheet.update_acell(category_cell, category)
        time.sleep(1)

    ### 인트로멘트 (J열) ###

    # J열 기존 데이터 확인
    ment_intro_cell = f"J{i}"
    existing_data = sheet.acell(ment_intro_cell).value

    if not existing_data : 
        # OpenAI API를 사용하여 열 J에 들어갈 내용 생성
        ment_intro_prompt = prompts.MENT_INTRO.format(content=body)  # 프롬프트 내 {body} 의 내용참조
        ment_intro_response = openai_client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": ment_intro_prompt}]
        )
        ment_intro = ment_intro_response.choices[0].message.content

        # 열 J에 내용 기록
        sheet.update_acell(ment_intro_cell, ment_intro)
        time.sleep(1)

    ### 카테고리 (K열) ###

    # F열 기존 데이터 확인
    closing_cell = f"K{i}"
    existing_data = sheet.acell(closing_cell).value

    if not existing_data : 
        # OpenAI API를 사용하여 열 K에 들어갈 내용 생성
        closing_prompt = prompts.MENT_CLOSING.format(content=body)  # 프롬프트 내 {body} 의 내용참조 
        closing_response = openai_client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": closing_prompt}]
        )
        closing = closing_response.choices[0].message.content

        # 열 K에 내용 기록
        sheet.update_acell(closing_cell, closing)
        time.sleep(1)


    # 이미지 키워드 단어 선택 및 검색 (G, H, I 열) 
    search_terms = sheet.cell(i, 6).value.split(',')
    body_content = sheet.cell(i, 4).value
    print(f"Row {i} [ 검색어 리스트 ] - {search_terms}")

    for j, term in enumerate(search_terms):
        if j >= 3: 
            break

        cleaned_term = term.strip().replace('#', '')

        # 이미지 웹사이트 검색 - 이미지정보 가져옴
        images_info = duckduckgo_image_search(cleaned_term)
        print(f"[ 이미지 검색결과 출력 ] - {cleaned_term} : {images_info}")


        # 이미지선택, URL추출
        selected_image_urls = analyze_images_with_openai(images_info, body_content)
        print(f"[ AI가선택한키워드 및 URL ] - {term} : {selected_image_urls}")


        # 이미지 시트에 저장
        for k, url in enumerate(selected_image_urls[:3]):  # 상위 3개 이미지만 저장
            cell = chr(ord('G') + k) + str(i)  # G, H, I 열에 해당하는 셀 주소 계산
            sheet.update_acell(cell, url)  # 이미지 URL 저장

            print(f"Updating cell {cell} with URL {url}")

            time.sleep(2)  # API 요청 사이에 지연 시간 추가