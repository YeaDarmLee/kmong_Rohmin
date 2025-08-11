import time, subprocess, re, math
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

subprocess.Popen('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\coupangCookie\\kmong_Rohmin"'.format("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"))

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
driver.get('https://nol.yanolja.com/entertainment/36052')

time.sleep(2)

# 전체 검색결과 수
text = driver.find_element(By.CLASS_NAME, "total-count").text
match = re.search(r'\d+', text)
total_count = int(match.group()) if match else 0
pages = math.ceil(total_count / 20)

# 현재 검색 수 초기화
current_count = 0

# 페이지 상세 URL 추출
print(" !! 페이지 상세 URL 추출 중 ... ")
imageHrefList = []
for page in range(1, pages + 1):
  time.sleep(1)
  
  imageLinkList = driver.find_elements(By.CLASS_NAME, "image-link")
  
  for imageLink in imageLinkList:
    imageHref = imageLink.get_attribute('href')
    imageHrefList.append(imageHref)
    current_count += 1
  
  if (page < 10):
    driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/main/div/section/div[2]/button[' + str(page + 2) + ']').click()
  elif (page == 10):
    driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/main/div/section/div[2]/button[12]').click()
  elif (page < pages):
    driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/main/div/section/div[2]/button[' + str(page - 8) + ']').click()

print(f" !! 페이지 상세 URL 추출 완료!! : 총 추출 개수 {len(imageHrefList)} 개")

# 추출한 상세페이지 URL 검색 및 엑셀 저장
results = []
idx = 1
for url in imageHrefList:
  print(f"{idx} : !! 상세 URL 접속 !! : 대상 {url} ")
  driver.get(url)
  time.sleep(1)
  # 공모전 명
  contestName = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/main/div/div/section[1]/header/h1').text
  # 주최자
  organizationName = driver.find_element(By.CLASS_NAME, "organization-name").text
  # 주최자 유형 (중앙정부/공기업/대기업 등)
  organizationType = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/section[1]/div/article/div[1]/dl[1]/dd").text
  # 신청자유형 (대학생/일반인 등)
  applicantType = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/section[1]/div/article/div[1]/dl[2]/dd").text
  # 썸네일 URL
  a_tag = driver.find_element(By.CLASS_NAME, 'card-image')
  thumbnailURL = a_tag.get_attribute('src')
  # 카테고리(영상)
  category = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/section[1]/div/article/div[1]/dl[7]/dd/ul").text
  # 시작일
  startDate = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/section[1]/div/article/div[1]/dl[4]/dd/div/span[2]").text
  # 마감일
  endDate = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/section[1]/div/article/div[1]/dl[4]/dd/span[2]").text
  # 총상금
  totalPrize = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/section[1]/div/article/div[1]/dl[3]/dd").text
  # 홈페이지 URL
  homepageURL = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/section[1]/div/article/div[1]/dl[5]/dd").text
  # 상세URL
  currentUrl = driver.current_url
  # 상세내용
  content = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/article/section[1]").get_attribute("outerHTML")

  # 엑셀
  results.append({
    "구분": idx,
    "공모전명": contestName,
    "주최자 유형": organizationType,
    "주최자": organizationName,
    "신청자 유형": applicantType,
    "썸네일 URL": thumbnailURL,
    "공모전카테고리": category,
    "시작일": startDate,
    "마감일": endDate,
    "총상금": totalPrize,
    "홈페이지 URL": homepageURL,
    "상세 URL": currentUrl,
    "content": content
  })

  idx += 1

# 엑셀 파일로 저장
df = pd.DataFrame(results)
df.to_excel("linkareer_c_result.xlsx", index=False, engine='openpyxl')