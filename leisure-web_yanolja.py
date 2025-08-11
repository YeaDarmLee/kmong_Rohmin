import time, subprocess, json, re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

subprocess.Popen('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\chromeCookie\\kmong_Rohmin_leisure"'.format("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"))

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
with open('leisure-web_yanolja.json', 'r', encoding='utf-8') as f:
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

    time.sleep(0.3)

    # aria-label="program-information" 내부 모든 img 태그 찾기
    imgs = driver.find_elements(By.CSS_SELECTOR, 'div[aria-label="program-information"] img')
    # src 속성만 추출
    image_urls = [img.get_attribute("src") for img in imgs]

    # 상품관련정보
    company_div = driver.find_element(By.CSS_SELECTOR, 'div.py-20')
    companyInfo = company_div.text

    # "더보기"가 있는 경우 버튼 클릭
    if "더보기" in companyInfo:
        try:
            # div 내부 버튼 클릭
            more_button = company_div.find_element(By.TAG_NAME, 'button')
            more_button.click()

            # 업데이트된 텍스트 다시 가져오기
            companyInfo = driver.find_element(By.CSS_SELECTOR, 'div.py-20').text
            
            # '접기'가 있으면 제거
            if companyInfo.endswith("접기"):
                companyInfo = companyInfo.rstrip("접기").strip()
        except Exception as e:
            print("더보기 버튼 클릭 실패:", e)

    # 스크롤 아래로 내리기
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    
    # 판매자정보 클릭
    driver.find_element(By.XPATH, '(//div[@class="py-20"])[last()]//button').click()
    time.sleep(0.3)
    
    # 상호
    corporationName = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[1]/td").text
    # 대표자명
    bossName = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[2]/td").text
    # 사업자등록번호
    registrationNumber = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[6]/td").text
    # E-mail
    email = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[4]/td").text
    # 연락처
    companyPhone = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[5]/td").text
    # 주소
    address = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[3]/td").text
    # 통신판매업
    orderBusiness = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[7]/td").text
    
    # 엑셀
    results.append({
      "상품관련정보": companyInfo,
      "상호명": corporationName,
      "대표자명": bossName,
      "사업자주소": address,
      "전자우편주소": email,
      "연락처": companyPhone,
      "사업자등록번호": registrationNumber,
      "통신판매업신고": orderBusiness,
      "랜딩페이지": image_urls
    })
  except:
    results.append({
      "상품관련정보": '에러 재확인 필요',
      "상호명": '에러 재확인 필요',
      "대표자명": '에러 재확인 필요',
      "사업자주소": '에러 재확인 필요',
      "전자우편주소": '에러 재확인 필요',
      "연락처": '에러 재확인 필요',
      "사업자등록번호": '에러 재확인 필요',
      "통신판매업신고": '에러 재확인 필요',
      "랜딩페이지": '에러 재확인 필요'
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
df.to_excel("leisure-web_yanolja_result.xlsx", index=False, engine='openpyxl')