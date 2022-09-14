import copy
import sys
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from bs4 import BeautifulSoup
import re

## 도로명 주소 -> 지번 주소 크롤링하기!
chrome_options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.implicitly_wait(10)
driver.get("https://www.juso.go.kr/support/AddressMainSearch.do?firstSort=none&ablYn=N&aotYn=N&fillterHiddenValue=&searchKeyword=&dsgubuntext=&dscity1text=&dscounty1text=&dsemd1text=&dsri1text=&dssan1text=&dsrd_nm1text=&searchType=HSTRY&dssearchType1=road&dscity1=&dscounty1=&dsrd_nm_idx1=%EA%B0%80_%EB%82%98&dsrd_nm1=&dsma=&dssb=&dstown1=&dsri1=&dsbun1=&dsbun2=&dstown2=&dsbuilding1=")

cmd = input(r"$ 불러올 엑셀 파일의 이름을 입력하세요. (예: 08. 인삼제품 생산정보 v1.0.xlsx) (실행파일과 같은 폴더에 있어야 합니다)"+"\n>>> ")
try:
    df = pd.read_excel("./"+cmd)
except FileNotFoundError:
    print("해당 파일을 찾을 수 없습니다. 프로그램을 종료합니다.")
except Exception as e:
    print("알 수 없는 오류입니다. 프로그램을 종료합니다. >>", e)

find2_index = ["소재지(지번)", "주소(지번)", "재배필지주소(지번)"]
find_index = ["소재지(도로명)", "주소(도로명)", "재배필지주소(도로명)"]
finded_data = pd.DataFrame()

for i in find_index:
   if i in df.columns:
        finded_data[i] = df[i]
        finded_data[find2_index[find_index.index(i)]] = df[find2_index[find_index.index(i)]]

# 검색창에 검색어 입력
idx = 0
for i in finded_data[finded_data.columns[0]]:
    if len(i) < 2:
        idx += 1
        continue
    driver.find_element(By.ID, "keyword").click()
    driver.find_element(By.ID, "keyword").clear()
    driver.find_element(By.ID, "keyword").send_keys(i)
    driver.find_element(By.XPATH, "//*[@id='searchButton']").click()
    time.sleep(1)
    html = driver.page_source
    soup = BeautifulSoup(html, "html.parser")
    if not soup.select("div.subject_area span.roadNameText"):
        idx += 1
        continue
    finded_data[finded_data.columns[1]][idx] = soup.select("div.subject_area span.roadNameText")[1].text.strip().replace("\n", "").replace("\t", "")
    idx += 1

df[finded_data.columns[1]] = finded_data[finded_data.columns[1]]
df.to_excel("./result.xlsx")

driver.close()
