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
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import subprocess
import shutil
import chromedriver_autoinstaller
from selenium.webdriver.chrome.options import Options



service = webdriver.ChromeService(service_args=['--disable-build-check'], log_output=subprocess.STDOUT)

# 크롬 드라이버 경로 지정
chrome_driver_path = r"C:\chrome-120.0.6099.109\120\chromedriver.exe"

# 새로운 Chrome 세션 시작
subprocess.Popen(
    r'C:\chrome-120.0.6099.109\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrometemp"')

# Chrome 옵션 설정
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

# ChromeDriver 자동 설치 및 실행
service = Service(executable_path=r'C:\chrome-120.0.6099.109\120\chromedriver.exe')
chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
try:
    driver = webdriver.Chrome(
        service = service, 
        options = chrome_options
    )
    driver.get("http://www.facebook.com/")
except:
    chromedriver_autoinstaller.install(True)
    driver = webdriver.Chrome(
        service = service, 
        options = chrome_options
    )
driver.implicitly_wait(10)



def duckduckgo_image_search(query):
    image_info = []  # 이미지 정보를 저장할 빈 리스트

    try:
        # DuckDuckGo 이미지 검색 URL로 바로 이동
        search_url = f"https://duckduckgo.com/?q={query}&ia=images&iax=images"
        driver.get(search_url)
        time.sleep(5)  # 페이지가 로드될 때까지 기다림

        # 검색 결과 중 첫 번째 이미지 선택
        first_image = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".tile--img__img.js-lazyload"))
        )
        first_image.click()
        time.sleep(5)  # 이미지가 로드될 때까지 기다림

        # 현재 페이지의 HTML 소스 가져오기 및 파싱
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')

        # 이미지 정보 추출
        images = soup.find_all('div', {'class': 'detail__media__img-wrapper'})
        for img_wrapper in images[:20]:
            img_link = img_wrapper.find('a', {'class': 'detail__media__img-link'})  # 클래스명 수정
            if not img_link:
                continue

            image_src = img_link.get('href')  # 원본 이미지 URL 추출

            img_highres = img_wrapper.find('img', {'class': 'detail__media__img-highres'})
            if img_highres:
                image_src = img_highres.get('src')  # 실제 이미지 URL
                image_title = img_highres.get('alt', '')  # 이미지 제목 수정

            # 파일명 추출
            filename = image_src.split('/')[-1] if image_src else ''

            # 이미지 크기 추출
            image_size_div = img_wrapper.find('div', {'class': 'c-detail__filemeta'})  # 클래스명 수정
            image_size = image_size_div.text if image_size_div else ''

            image_info.append({
                'filename': filename,
                'url': image_src,
                'title': image_title,
                'size': image_size
            })

    except Exception as e:
        print(f"오류 발생: {e}")

    finally:
        driver.quit()

    return image_info


# 이미지 첫 번째 프롬프트 : 웹에 검색할 단어 선택(검색하는 데 까지 관여)

def select_relevant_images(images, content):
    prompt = prompts.IMAGE_TEMPLATE_1.format(content=content)
    for i, img in enumerate(images):
        prompt += f"{i+1}: URL: {img}\n"

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "system", "content": prompt}],
        max_tokens=150  # 토큰 수를 증가시켜 여러 이미지를 선택할 수 있도록 함
    )

    # ChatGPT의 응답에서 세 개의 이미지 URL 추출
    response_text = response.choices[0].message['content'].strip().split('\n')
    selected_urls = [url for url in response_text if url.startswith('http')]
    return selected_urls[:3]  # 상위 세 개의 URL 반환


# 이미지2번째 프롬프트 : 이미지 정보를 OpenAI API를 사용하여 본문글과 비교,
# 선택하여 스프레드 시트로 가져옵니다. 
def analyze_images_with_openai(image_info, content):
    prompt = prompts.IMAGE_TEMPLATE_2.format(content=content)
    for i, img in enumerate(image_info):
        prompt += f"{i+1}: 파일명: {img['filename']}, URL: {img['url']}, 제목: {img['title']}, 사이즈: {img['size']}px\n"

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "system", "content": prompt}],
        max_tokens=150
    )

    return response.choices[0].message['content'].strip().split('\n')
    print(f' ')


# 마지막 데이터가 있는 행 찾기
def find_last_row(sheet):
    str_list = filter(None, sheet.col_values(2))  # 2번째 열(B열)의 값들을 가져옴
    return len(list(str_list)) + 1  # 마지막 행 번호 반환



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

row = find_last_row(sheet)  # 시작 행을 마지막 행 다음으로 설정
titles_to_create = []  # 생성할 제목을 저장할 리스트

# 제목 목록 생성 (한 번만 수행)
if not titles_to_create:
    title_prompt = prompts.TITLE_PROMPT_TEMPLATE.format(main_concept=main_concept, sub_concept=sub_concept)
    title_response = openai.Chat.Completion.create(
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
        body_response = client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": body_prompt}]
        )
        body = body_response.choices[0].message['content'].strip()
        
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
        summary_response = client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": summary_prompt}]
        )
        summary = summary_response.choices[0].message['content'].strip()

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
        tag_response = client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": tag_prompt}]
        )
        tag = tag_response.choices[0].message['content'].strip()

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
        category_prompt = prompts.CATEGORY_PROMPT_TEMPLATE.format(content=body)
        category_response = client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": category_prompt}]
        )
        category = category_response.choices[0].message['content'].strip()

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
        ment_intro_response = client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": ment_intro_prompt}]
        )
        ment_intro = ment_intro_response.choices[0].message['content']

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
        closing_response = client.chat.completions.create(
            model='gpt-3.5-turbo',
            messages=[{"role": "system", "content": closing_prompt}]
        )
        closing = closing_response.choices[0].message['content']

        # 열 K에 내용 기록
        sheet.update_acell(closing_cell, closing)
        time.sleep(1)




    # 이미지 키워드 단어 선택 및 검색 (G, H, I 열) 
    search_terms = sheet.cell(i, 6).value.split(',')
    content = sheet.cell(i, 4).value
    print(f"Row {i} [ 검색어 리스트 ] - {search_terms}")
    all_images_info = []

    for j, term in enumerate(search_terms):
        if j >= 3: 
            break

        cleaned_term = term.strip().replace('#', '')
        images_info = duckduckgo_image_search(cleaned_term)
        all_images_info.extend(images_info)
        print(f"[ 이미지 검색결과 출력 ] - {cleaned_term} : {images_info}")


        # 이미지선택, URL추출
        selected_image_urls = analyze_images_with_openai(images_info, content)
        print(f"[ AI가선택한키워드 및 URL ] - {term} : {selected_image_urls}")


        # 이미지 시트에 저장
        for k, url in enumerate(selected_image_urls[:3]):  # 상위 3개 이미지만 저장
            cell = chr(ord('G') + k) + str(i)  # G, H, I 열에 해당하는 셀 주소 계산
            sheet.update_acell(cell, url)  # 이미지 URL 저장

            print(f"Updating cell {cell} with URL {url}")

            time.sleep(2)  # API 요청 사이에 지연 시간 추가


        
 

     
 


