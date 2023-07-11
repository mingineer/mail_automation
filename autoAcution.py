# 4버전 셀레니움 사용 시 크롬 드라이버 자동 다운
#from selenium.webdriver.chrome.service import Service
#from selenium.webdriver.chrome import ChromeDriverManager
import datetime
from multiprocessing.connection import wait
import os
import time
import win32com.client
from datetime import datetime
from os import listdir
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
import io
import base64
from PIL import Image
import openpyxl as op
from selenium.webdriver.common.alert import Alert

# 이미지 자르기
def crop_image(png, x, y, xr, yr,img_name):
   
   image1 = Image.open(png)
   print(image1.size)
   croppedImage=image1.crop((x, y, xr, yr))
   print("잘려진 사진 크기: ", croppedImage.size)

   croppedImage.save(img_name)

# 메일 자동 발송
def send_mail(to,cc,subject,content, atch=[]):
    new_Mail = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
    new_Mail.To = to
    new_Mail.CC = cc
    new_Mail.Subject = subject
    new_Mail.HTMLBody = content

    # 첨부파일 추가
    #if atch:
    #    for file in atch:
    #        new_Mail.Attachments.Add(file)
    new_Mail.Send()

#이미지 변환 함수
def chg_jpg(png):
    global bytearr
    global rgb_im
    global imgbytearr
    global encoded_image
    global image_for_body

    img = Image.open(png)
    # 투명도 지원 설정
    rgb_im = img.convert('RGB')

    bytearr = io.BytesIO()
    rgb_im.save(bytearr, format="JPEG")
    
    imgbytearr = bytearr.getvalue()

    encoded_image = base64.b64encode(imgbytearr).decode("utf-8")
    image_for_body = f'<img src="data:image/png;base64,{encoded_image}"/>'





today = datetime.today()
today_send = today.strftime("%m/%d")

options = webdriver.ChromeOptions()
options.add_argument('--log-level=3')
# options.add_argument('headless')

# 4버전 사용 시
# browser = webdriver.Chrome(service=Service(ChromeDriverManager().install))
browser = webdriver.Chrome('./chromedriver.exe', options=options)


browser.get('${BROWSER_URL}')
browser.maximize_window() # 창 최대화
time.sleep(3)

# 로그인
print(datetime.now().strftime('%Y.%m.%d - %H:%M:%S'))

print('loging in')
browser.find_element(By.ID, 'user_id').send_keys('${ID}')
browser.find_element(By.ID, 'user_pw').send_keys('${PW}')
browser.find_element(By.XPATH, '//*[@id="doLogin"]').click()
print('login clicked')
time.sleep(5)


# 리소스 모니터링 페이지로 이동
browser.get('${MONITORING_PAGE}')
time.sleep(5)


alert = Alert(browser)
alert.accept()

# Alert 엔터 처리
#try:
#    sleep(5)
#    result = browser.switch_to_alert()
#    result.accept()
#    result.dismiss()

#except:
#    "There is no alert"

time.sleep(10)


# 서비스 선택
browser.switch_to.default_content() 
elem = browser.find_element(By.XPATH, '${SERVICE_PAGE_XPATH}')
elem.click()


# 대기시간 설정 (랜더링 대기)
time.sleep(300)



# 캡처 (element 캡처)
server_equip = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
server_equip_png = server_equip.screenshot_as_png  
with open("server_status.png", "wb") as file:  
    file.write(server_equip_png)

# 자르기
crop_image("server_status.png", 0, 10, 1490, 440, 'server_status.png')



# 검색
browser.find_element(By.ID, '${SEARCH_ID}').send_keys('${KEYWORD}')
time.sleep(5)

# 선택 
elem = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
elem.click()
time.sleep(5)

# Alert 엔터 처리 (조회된 데이터가 없습니다)
try:
    result = browser.switch_to_alert()
    result.accept()
    result.dismiss()

except:
    "There is no alert"

time.sleep(300)

# 캡처
svc_equip = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
svc_equip_png = svc_equip.screenshot_as_png  
with open("svc_status.png", "wb") as file:  
    file.write(svc_equip_png)

# 자르기
crop_image("svc_status.png", 0, 10, 1490, 65, 'svc_status.png')



# 이벤트 상태 확인
browser.get('${EVENT_PAGE}')
time.sleep(10)

# 전용 검색
event_svc_button = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
event_svc_button.click()

event_svc = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
event_svc.send_keys('전용') 

# 버튼 선택 
click_button = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
click_button.click()
time.sleep(2)

# 조회
Lookup = browser.find_element(By.ID, 'btn_search')
Lookup.click()
time.sleep(100)

# 캡처
indepen_ext = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
indepen_ext_png = indepen_ext.screenshot_as_png  
with open("이벤트.png", "wb") as file:  
    file.write(indepen_ext_png)

# 자르기
crop_image("이벤트.png", 0, 10, 1850, 255, '이벤트.png')
time.sleep(5)

# 새로고침 
browser.get('${MONITORING_PAGE}')
time.sleep(10)

# 검색
event_svc_button = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
event_svc_button.click()

event_svc = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
event_svc.send_keys('${SEARCH_SVC}') 


# 버튼 선택
click_button = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
click_button.click()
time.sleep(2)

# 조회
element4 = browser.find_element(By.ID, 'btn_search')
element4.click()
time.sleep(100)


# 캡처
svc_equip = browser.find_element(By.XPATH, '${RESOURCE_XPATH}')
svc_equip_png = svc_equip.screenshot_as_png  
with open("이벤트_1.png", "wb") as file:  
    file.write(svc_equip_png)

# 자르기
crop_image("이벤트_1.png", 0, 10, 1875, 255, '이벤트_1.png')




# 메일 본문 작성
contents = "안녕하세요, ~ <br><br> 서버 리소스 현황 <br> <br> <b> ~ </b> <br>"
auction = chg_jpg("server_status.png")
contents += image_for_body 
rpaawbs1 = chg_jpg("svc_status.png")
contents += image_for_body
ext_auction_evnt = chg_jpg("이벤트.png")
contents += "<br> <b> 현재 이벤트 상태<br><br>- ~ </b> <br>"
contents += image_for_body
contents += "<br> <b> - ~ </b> <br>"
auction_evnt= chg_jpg("이벤트_1.png")
contents += image_for_body
contents += "<br> <b> 이상입니다. </b> <br> <b> 감사합니다. </b> <br> <br><br><br> <b> 본 메일은 자동 발송된 메일입니다. 회신하지 마십시오. </b>"



# 메일 발송
print('메일 발송')
os.system('taskkill /im outlook.exe /f')

receivers = '${RECEIVERS}'
cc_users = '${CC}'

send_mail(receivers, cc_users, today_send + "~", contents)
print('메일 발송 완료')
browser.quit()

