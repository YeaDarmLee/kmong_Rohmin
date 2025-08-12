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

# JSON 문자열을 파일로 저장했다고 가정
with open('nol_yanolja_error.json', 'r', encoding='utf-8') as f:
  data = json.load(f)

# urls 값 가져오기
urls = data.get("urls", [])

# 추출한 데이터 엑셀 저장 변수
results = []
idx = 1

# 반복문 실행
for url in urls:
  print(f"{idx} : !! URL 접속 !! : {url} ")
  
  try:  
    # 사이트 진입
    driver.get(url)

    time.sleep(0.5)

    # iframe # iframe 전환
    iframe = driver.find_element(By.CSS_SELECTOR, 'iframe[title="iframe"]')
    driver.switch_to.frame(iframe)

    # 예매안내 팝업 닫기
    try:
      driver.find_element(By.XPATH, '/html/body/div[2]/div[4]/div/div[3]/button').click()
    except:
      time.sleep(0.1)

    time.sleep(0.5)
    
    # 타이틀
    title = driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div/div[2]/div/h2").text

    # 첫 번째 이미지 URL
    try:
      productsNotice = driver.find_element(By.CSS_SELECTOR, ".productsNotice")
      imgs = productsNotice.find_elements(By.TAG_NAME, "img")
      image_urls = [img.get_attribute("src") for img in imgs]
    except:
      image_urls = []

    # 두 번째 이미지 URL
    try:
      productsPriceInformation = driver.find_element(By.CSS_SELECTOR, ".productsPriceInformation")
      imgs1 = productsPriceInformation.find_elements(By.TAG_NAME, "img")
      image_urls1 = [img.get_attribute("src") for img in imgs1]
    except:
      image_urls1 = []

    # 세 번째 이미지 URL
    try:
      productsDetail = driver.find_element(By.CSS_SELECTOR, ".productsDetail")
      imgs2 = productsDetail.find_elements(By.TAG_NAME, "img")
      image_urls2 = [img.get_attribute("src") for img in imgs2]
    except:
      image_urls2 = []

    # 세 개 리스트 합치기
    all_image_urls = image_urls + image_urls1 + image_urls2

      # 엑셀
    results.append({
        "타이틀": title,
        "page": url,
        "상품관련정보": '-',
        "상호명": '정보 없음',
        "대표자명": '정보 없음',
        "사업자주소": '정보 없음',
        "전자우편주소": '정보 없음',
        "연락처": '정보 없음',
        "사업자등록번호": '정보 없음',
        "통신판매업신고": '-',
        "랜딩페이지": all_image_urls
    })
  except:
    results.append({
      "타이틀": '',
      "page": url,
      "상품관련정보": '에러 재확인 필요',
      "상호명": '-',
      "대표자명": '-',
      "사업자주소": '-',
      "전자우편주소": '-',
      "연락처": '-',
      "사업자등록번호": '-',
      "통신판매업신고": '-',
      "랜딩페이지": '-'
    })

  idx += 1
  
# 엑셀 파일로 저장
# 1) 결과를 DF로 만들기
df = pd.DataFrame(results)

# 2) 엑셀에서 금지되는 제어문자 제거 함수
_illegal = re.compile(r'[\x00-\x1F\x7F]')

def remove_illegal_chars(v):
  if isinstance(v, str):
    return _illegal.sub('', v)
  return v

# 3) 모든 셀에 적용
df = df.applymap(remove_illegal_chars)

# 4) 엑셀로 저장 (openpyxl)
df.to_excel("nol_yanolja_error_result.xlsx", index=False, engine='openpyxl')