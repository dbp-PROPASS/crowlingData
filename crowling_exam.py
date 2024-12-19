import requests 
import json
import time
from datetime import datetime
import random
import re
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook, load_workbook

# 크롬 드라이버 경로 및 URL
driver_path = "C:/Users/tlsgo/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe"
URL = 'https://www.q-net.or.kr/crf005.do?id=crf00501&gSite=Q&gId=#none'

# 크롬 드라이버 사용
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

#####################################cert_id 리스트 생성성 #################################################

# 기관 코드와 카테고리 매핑
org_categories = {
    'P200': ['10', '02'],
    'R139': ['08'],
    'N004': ['15'],
    'P317': ['21'],
    'R121': ['08'],
    'N003': ['21'],
    'R020': ['26'],
    'N002': ['21'],
}

# 자격증 ID 저장 리스트
# 자격증 ID와 이름 저장 리스트
cert_list = []

# 모든 org와 카테고리를 조합하여 반복
for org, categories in org_categories.items():
    for category in categories:
        time.sleep(random.uniform(2, 5))  # 요청 간 랜덤 대기 시간
        url = f'https://www.q-net.or.kr/crf005.do?id=crf00501s03&gSite=Q&gId=&obligFldCd={category}&examInstiCd={org}'

        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            response = requests.get(url, headers=headers, timeout=10)
            soup = BeautifulSoup(response.content, 'html.parser')

            for input_tag in soup.find_all('input', {'name': 'jmcd'}):
                jmCd = input_tag.get('value')
                if jmCd:
                    detail_url = f'https://www.q-net.or.kr/crf005.do?id=crf00501s04&mdobligFldCd={jmCd}&examInstiCd={org}'

                    for attempt in range(3):  # 최대 3번 재시도
                        try:
                            detail_response = requests.get(detail_url, headers=headers, timeout=10)
                            detail_soup = BeautifulSoup(detail_response.content, 'html.parser')

                            # CERT_ID와 CERT_NAME 추출
                            for a_tag in detail_soup.find_all('a', onclick=True):
                                match = re.search(r"jmDetail2\('(\d+)',\s*'([^']+)'\)", a_tag['onclick'])
                                if match:
                                    cert_id = match.group(1)  # 자격증 ID
                                    cert_name = match.group(2)  # 자격증 이름
                                    cert_list.append({'id': cert_id, 'name': cert_name})
                                    print(f"[SUCCESS] 자격증 ID: {cert_id}, 이름: {cert_name} 저장 완료")
                            break  # 성공하면 루프 종료
                        except requests.exceptions.RequestException as e:
                            print(f"[ERROR] 재시도 중 ({attempt + 1}/3): {e}")
                            time.sleep(random.uniform(3, 7))  # 재시도 간 대기 시간
                    else:
                        print(f"[ERROR] Detail URL 요청 실패: {detail_url}")
        except Exception as e:
            print(f"Error fetching data for {org}, category {category}: {e}")

# 최종 리스트 확인
print(cert_list)

# 산업인력공단 자격증 ID를 별도의 리스트로 저장
industry_cert_id_list = []

# Selenium으로 자격증 ID를 추출하는 코드
for idx in range(1, 27):
    obligFldCd = str(format(idx, '02'))  # 1 -> '01', 2 -> '02', ...
    url = f'https://www.q-net.or.kr/crf005.do?id=crf00501s01&gSite=Q&gId=&div=1&obligFldCd={obligFldCd}'
    
    # Selenium으로 URL 접속
    driver.get(url)
    time.sleep(2)  # 페이지 로딩 대기
    
    # 현재 페이지의 HTML을 BeautifulSoup으로 파싱
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    
    # 분야 ID 추출
    jmCd_ids = []  # 분야 ID를 저장할 리스트
    for a_tag in soup.find_all('a', onclick=True):
        match = re.search(r"step3BunRyu\('([0-9\-]+)',", a_tag['onclick'])
        if match:
            jmCd_id = match.group(1)  # 분야 ID
            jmCd_ids.append(jmCd_id)  # 리스트에 추가
    
    # 자격증 ID 추출
    for jmCd in jmCd_ids:
        detail_url = f'https://www.q-net.or.kr/crf005.do?id=crf00501p02&gSite=Q&gId=&jmCd={jmCd}&examInstiCd='
        
        # Selenium으로 상세 URL 접속
        driver.get(detail_url)
        time.sleep(2)  # 페이지 로딩 대기
        
        # 상세 페이지의 HTML을 BeautifulSoup으로 파싱
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        
        # 자격증 ID만 추출하여 리스트에 추가
        for a_tag in soup.find_all('a', onclick=True):
            match = re.search(r"jmDetail\('(\d+)',\s*'([^']+)'\);", a_tag['onclick'])
            if match:
                cert_id = match.group(1)  # 자격증 ID
                if cert_id not in industry_cert_id_list:  # 중복 확인
                    industry_cert_id_list.append(cert_id)  # 리스트에 추가

    # 중간에 자격증 ID 리스트를 JSON으로 저장하여 과부하 방지
    with open('industry_cert_ids.json', 'w', encoding='utf-8') as json_file:
        json.dump(industry_cert_id_list, json_file, ensure_ascii=False, indent=4)

# 산업인력공단 자격증 ID 리스트를 최종적으로 JSON으로 저장
output_file_industry = "industry_cert_ids.json"
with open(output_file_industry, "w", encoding="utf-8") as json_file:
    json.dump(industry_cert_id_list, json_file, ensure_ascii=False, indent=4)
print(f"[INFO] 산업인력공단 자격증 ID 리스트가 {output_file_industry} 파일로 저장되었습니다.")


# 다른 기관의 자격증 ID 리스트를 JSON으로 저장
output_file_other = "other_cert_ids.json"
with open(output_file_other, "w", encoding="utf-8") as json_file:
    json.dump(cert_list, json_file, ensure_ascii=False, indent=4)
print(f"[INFO] 다른 기관의 자격증 ID 리스트가 {output_file_other} 파일로 저장되었습니다.")

#####################################산업인력공단 일정 스크래핑#################################################
#####round_id#######

# cert_id 리스트 로드
with open('industry_cert_ids.json', 'r', encoding='utf-8') as json_file:
    cert_id_list = json.load(json_file)


exam_info_list = []

# 날짜 추출 함수
def extract_dates_from_text(text):
    # HTML 태그 및 주석 제거
    cleaned_text = re.sub(r'<.*?>|<!--.*?-->', '', text, flags=re.DOTALL)
    cleaned_text = re.sub(r'\[.*?\]', '', cleaned_text)  # 빈자리 접수 등 제거
    cleaned_text = re.sub(r'\s+', '', cleaned_text)  # 공백 제거
    
    # 날짜 구간 추출
    date_pairs = re.findall(r'(\d{4}\.\d{2}\.\d{2})', cleaned_text)
    
    # 날짜 포맷 변환 함수
    def format_date(date_str):
        try:
            return datetime.strptime(date_str, "%Y.%m.%d").strftime("%Y/%m/%d")
        except ValueError:
            return None  # 날짜 형식이 잘못되었을 경우 None 반환

    if len(date_pairs) >= 2:
        return format_date(date_pairs[0]), format_date(date_pairs[1])  # 시작일, 종료일
    elif len(date_pairs) == 1:
        return format_date(date_pairs[0]), None  # 시작일만 있는 경우
    return None, None  # 날짜가 없을 경우


# 각 cert_id에 대해 round_id 추출
for cert_id in cert_id_list:
    # 상세 URL 설정
    detail_url = f'https://www.q-net.or.kr/crf005.do?id=crf00503s02&gSite=Q&gId=&jmCd={cert_id}'
    
    # Selenium으로 URL 접속
    driver.get(detail_url)
    # time.sleep(2)  # 페이지 로딩 대기
    
    # BeautifulSoup으로 HTML 파싱
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # 시험 일정에서 round_id 추출
    try:
        exam_schedule_div = soup.find('div', class_='contTable typeSm')
        if exam_schedule_div:
            exam_schedule_table = exam_schedule_div.find('table')
            if exam_schedule_table:
                rows = exam_schedule_table.find_all('tr')

                for row in rows:
                    columns = row.find_all('td')
                    round_name_tag = row.find('th')  # 회차 이름 (정기 기사 1회 등)
                    
                    # round_id 추출
                    round_id = None
                    if len(columns) > 1:
                        # 각 열에서 '2024년 수시 기사 1회'와 같은 텍스트를 찾기
                        round_name_match = re.search(r"(\d{4}년 \S+ \S+ \d+회)", columns[0].text.strip())
                        if round_name_match:
                            round_name = round_name_match.group(1)
                            # round_name에서 연도와 회차 추출
                            round_id_match = re.search(r"(\d{4})\D+(\d+)", round_name)
                            if round_id_match:
                                year = round_id_match.group(1)  # 2024년 -> 2024
                                round_number = round_id_match.group(2).zfill(2)  # 회차가 1자리일 경우 2자리로 맞춤 (1 -> 01)
                                round_id = f"{cert_id}{year[2:]}{round_number}"  # cert_id + 연도 2자리 + 회차 2자리 -> 12342401
                                # print(f"Extracted round_id: {round_id}")  # 디버깅용 출력

                    if len(columns) >= 6:  # 필기 & 실기 열이 있는 경우
                        # 필기 일정이 있는지 확인
                        if columns[1].text.strip():  # 필기 접수 시작일이 있는 경우
                            registration_start, registration_end = extract_dates_from_text(columns[1].decode_contents())
                            exam_start, exam_end = extract_dates_from_text(columns[2].decode_contents())
                            result_announcement, _ = extract_dates_from_text(columns[3].decode_contents())

                            exam_info_list.append([
                                f"{round_id}" if round_id else None,
                                registration_start,
                                registration_end,
                                result_announcement,
                                exam_start,
                                exam_end,
                                cert_id,
                                '필기'
                            ])

                        # 실기 일정이 있는지 확인
                        if columns[4].text.strip():  # 실기 접수 시작일이 있는 경우
                            registration_start, registration_end = extract_dates_from_text(columns[4].decode_contents())
                            exam_start, exam_end = extract_dates_from_text(columns[5].decode_contents())
                            result_announcement, _ = extract_dates_from_text(columns[6].decode_contents())

                            exam_info_list.append([
                                f"{round_id}" if round_id else None,
                                registration_start,
                                registration_end,
                                result_announcement,
                                exam_start,
                                exam_end,
                                cert_id,
                                '실기'
                            ])
    except Exception as e:
        print(f"Error processing cert_id {cert_id}: {e}")

# 결과 확인
# for exam in exam_info_list:
#     print(exam)

# print(len(exam_info_list))

excel_filename = "examSchedule.xlsx"
wb = Workbook()  # 새 워크북 생성
ws = wb.active
ws.title = "examSchedule"

# 헤더 추가
header = ['ROUND ID', '접수시작일', '접수마감일', '결과발표일', '시험시작일', '시험마감일', 'CERT_ID', '시험타입']
ws.append(header)

# 첫 번째 데이터 추가
for row in exam_info_list:
    ws.append(row)

# 엑셀 저장
wb.save(excel_filename)
print(f"[INFO] '{excel_filename}' 파일 생성 및 첫 번째 데이터 저장 완료.")

# 크롬 드라이버 종료 대기
print("[INFO] 크롬 드라이버를 종료하지 않고 유지합니다. 수동으로 종료할 때까지 유지됩니다.")
time.sleep(60*5)  # 5분간 대기
