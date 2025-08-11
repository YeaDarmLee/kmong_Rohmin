import time, subprocess, json, re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

subprocess.Popen('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\chromeCookie\\kmong_Rohmin_nol"'.format("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"))

# Selenium 옵션 설정
options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

# ChromeDriver 실행
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


# 사이트 진입
driver.get('https://nol.yanolja.com/entertainment/37602')

time.sleep(0.3)

# iframe 전환
iframe = driver.find_element(By.CSS_SELECTOR, 'iframe[title="iframe"]')
driver.switch_to.frame(iframe)

# 예매안내 팝업 닫기
try:
  driver.find_element(By.XPATH, '/html/body/div[2]/div[4]/div/div[3]/button').click()
except:
  time.sleep(0.1)

# 타이틀
title = driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div/div[2]/div/h2").text

# 첫 번째 이미지 URL
productsPriceInformation = driver.find_element(By.CSS_SELECTOR, ".productsPriceInformation")
imgs = productsPriceInformation.find_elements(By.TAG_NAME, "img")
image_urls = [img.get_attribute("src") for img in imgs]

# 두 번째 이미지 URL
productsDetail = driver.find_element(By.CSS_SELECTOR, ".productsDetail")
imgs2 = productsDetail.find_elements(By.TAG_NAME, "img")
image_urls2 = [img.get_attribute("src") for img in imgs2]

# 두 리스트 합치기
image_urls.extend(image_urls2)  # image_urls 뒤에 image_urls2 내용을 붙임

# 판매정보 탭 클릭
driver.find_element(By.CLASS_NAME, "productsTabAdditional").click()
time.sleep(0.3)

# 상품관련정보
try:
  companyInfo = driver.find_element(By.CLASS_NAME, "companyInfo").text
except:
  companyInfo = '-'

# 판매자정보 클릭
driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/div[4]/div[2]/div/div[1]').click()
# 상호
corporationName = driver.find_element(By.ID, "corporationName").text
# 대표자명
bossName = driver.find_element(By.ID, "bossName").text
# 사업자등록번호
registrationNumber = driver.find_element(By.ID, "registrationNumber").text
# E-mail
email = driver.find_element(By.ID, "email").text
# 연락처
companyPhone = driver.find_element(By.ID, "companyPhone").text
# 주소
address = driver.find_element(By.ID, "address").text

# 엑셀
print({
  "타이틀": title,
  "상품관련정보": companyInfo,
  "상호명": corporationName,
  "대표자명": bossName,
  "사업자주소": address,
  "전자우편주소": email,
  "연락처": companyPhone,
  "사업자등록번호": registrationNumber,
  "통신판매업신고": '-',
  "랜딩페이지": image_urls
})