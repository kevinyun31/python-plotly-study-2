import requests
from bs4 import BeautifulSoup
import openpyxl

fpath = r'C:\workspace\python_lab\STARTCODING\01_네이버_주식현재가_크롤링\data.xlsx'
wb = openpyxl.load_workbook(fpath)
ws = wb.active  # 현재 활성화된 시트 선택

# 종목 코드 리스트
codes = [
    '005930',  # 삼성전자
    '000660',  # SK하이닉스
    '035720',  # 카카오
    '080220',  # 제주반도체
    '005490',  # POSCO홀딩스
    '068270',  # 셀트리온
    '035420',  # NAVER
    '022100',  # 포스코DX
    '086520',  # 에코프로
    '047050'   # 포스코인터내셔널     
]

row = 2 
for code in codes:
    url = f"https://finance.naver.com/item/sise.naver?code={code}"
    response = requests.get(url)
    html = response.text
    soup = BeautifulSoup(html, 'html.parser')
    price = soup.select_one("#_nowVal").text
    price = price.replace(',', '')
    print(price)
    ws[f'B{row}'] = int(price)
    row = row + 1     
    
wb.save(fpath)    
print(" <모든 작업이 성공적으로 완료되었습니다> ")