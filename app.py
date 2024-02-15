# #-*- coding: utf-8 -*-
# ### 2024.01.29 
# ### 작성자 : 유정원

# pip install selenium
# pip install webdriver-manager
# # pip install schedule
# # pip install pymsteams
# pip install configparser

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# from openpyxl import Workbook
# from datetime import datetime

# import os
import configparser
import time
# import pymsteams
# import schedule
# import fileUploadTest

## Read setting.ini
properties = configparser.ConfigParser()
properties.read('./setting.ini')

EV6 = properties["EV6"]
Spotage = properties["Spotage"]

## selenuim Chrome Driver Option Setting
options = webdriver.ChromeOptions()
options.add_argument('lang=ko_KR') # 사용언어 한국어
# options.add_argument('disable-gpu') # 하드웨어 가속 안함
options.add_argument('log-level=3')
options.add_argument('--window-size=1920,1080') # 해상도 설정 
options.add_argument('headless')
# options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
options.add_experimental_option("detach",True) # 웹브라우저 유지

service = Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)



URL = Spotage["URL"]
driver.get(URL)

driver.find_element(By.CLASS_NAME, 'conf-btn-close').click()
driver.implicitly_wait(3)

## 트림 선택
trim = str("//label[@for='" + Spotage["TRIM"] + "']")
driver.find_element(By.XPATH, trim).click()
# driver.find_element(By.XPATH, "//label[@for='filterItem_0_NQ-9']").click()
driver.implicitly_wait(3)

## 트림 세부
trimItem = str("//label[@for='" + Spotage["TRIM_ITEM"] + "']")
driver.find_element(By.XPATH, trimItem).click()
driver.implicitly_wait(3)

## 컬러 선택 이동
driver.find_element(By.XPATH, '//*[@id="build_menu_btns"]/div/div/div/div/button').click()
driver.implicitly_wait(3)

## 외장 컬러 선택
extColor = str("//label[@for='" + Spotage["EXT_COLOR"] + "']")
driver.find_element(By.XPATH, extColor).click()
driver.implicitly_wait(3)

## 내장 컬러 선택
intColor = str("//label[@for='" + Spotage["INT_COLOR"] + "']")
driver.find_element(By.XPATH, intColor).click()
driver.implicitly_wait(3)

## 옵션 선택
driver.find_element(By.XPATH, '//*[@id="build_menu_btns"]/div/div/div/div[2]/button').click()
driver.implicitly_wait(3)

## 패키지 선택
package = str(Spotage["PACKAGE"])
if package == "none":
    pass
else:
    print("code exec")

## 상세 옵션 선택
option = str("//label[@for='" + Spotage["OPTION"] + "']")
driver.find_element(By.XPATH, option).click()
driver.implicitly_wait(3)

## 액세사리 선택
driver.find_element(By.XPATH, '//*[@id="build_menu_btns"]/div/div/div/div[2]/button').click()
driver.implicitly_wait(3)

## 견적 완료
driver.find_element(By.XPATH, '//*[@id="build_menu_btns"]/div/div/div/div[2]/button').click()
driver.implicitly_wait(5)

## 면세 구분 클릭
button_element = driver.find_element(By.XPATH, '//button[@data-link-label="면세 구분 및 등록비_상세보기"]')
driver.execute_script("arguments[0].click();", button_element)

### 일반인
paymentPrice = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div/div[4]/h3/button/span[2]/span[1]').text
paymentPrice = paymentPrice[:-1]
paymentPrice = paymentPrice.replace(",", "")
print("결제 금액 : " + paymentPrice)


### 일반인 + 다자녀

### 장애 4~6 등급

### 국가유공자

### 광주민주화





# 드롭다운
# select_element = driver.find_element(By.CLASS_NAME, 'dropdown-select__list')
# select = Select(select_element)
# select.select_by_value("장애 1~3급")
# driver.find_element(By.XPATH, '//*[@id="accordion-item-1707963078660-panel"]/div[1]/div[2]/div[1]/div/ul/li[2]/button').click()






### 확인 필요한 것들
# - 디폴트 사양?

time.sleep(5)
# conf-btn-close







# def main():
#     ## Read setting.ini
#     properties = configparser.ConfigParser()
#     properties.read('./setting.ini')
#     kia = properties["Kia"]
#     default = properties["Default"]
#     teams_key = properties["Default"]["teamsWebhook"]

#     ## Teams setting
#     teamsMsg = pymsteams.connectorcard(teams_key)

#     ## 전역변수 
#     global URL, ID, PW, CURRENT_PATH, RESULT_PATH

#     URL = "https://" + kia["url"]
#     ID = kia["id"]
#     PW = kia["pw"]
#     CURRENT_PATH = os.getcwd()
#     RESULT_PATH = os.path.join(CURRENT_PATH, "result/")

#     ## 테스트 결과 세팅(Teams)
#     teams_section = pymsteams.cardsection()
    
#     currentTime = str(datetime.today().strftime("%Y%m%d_%H%M%S"))
#     print("(Page)테스트 시작 시간 : " + currentTime)

#     ## selenuim Chrome Driver Option Setting
#     options = webdriver.ChromeOptions()
#     options.add_argument('lang=ko_KR') # 사용언어 한국어
#     # options.add_argument('disable-gpu') # 하드웨어 가속 안함
#     options.add_argument('headless')
#     # options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
#     # options.add_experimental_option("detach",True) # 웹브라우저 유지

#     service = Service(executable_path=ChromeDriverManager().install())
#     driver = webdriver.Chrome(service=service, options=options)

#     ## 메인 페이지 호출
#     start_time_Entity01 = time.time()
#     driver.get(URL)
#     end_time_Entity01 = time.time()
#     entity01 = end_time_Entity01 - start_time_Entity01
#     print("Login Page : " + str(entity01))
#     teams_section.addFact("Login Page", str(entity01))

#     ## ID/PW 입력 & 로그인
#     driver.find_element(By.NAME, 'userName').send_keys(ID)
#     driver.find_element(By.NAME, 'password').send_keys(PW)
#     driver.find_element(By.CLASS_NAME, 'btn_login').click()

#     start_time_Entity02 = time.time()

#     ## Alert창 처리(로그인 시간 확인)
#     try:
#         WebDriverWait(driver, 5).until(EC.alert_is_present())    
        
#         alert1 = driver.switch_to.alert
#         # 경고창 문구 출력
#         print("--------Alert--------")
#         print(alert1.text)
#         print("---------------------")

#         # 경고창 처리
#         alert1.accept()
#         time.sleep(0.5)
#     except:    
#         pass

#     end_time_Entity02 = time.time()
#     entity02 = end_time_Entity02 - start_time_Entity02 - 0.5
#     print("Login Page>Main Page : " + str(entity02))
#     teams_section.addFact("Login Page>Main Page", str(entity02))

#     ## By CATEGORY>All Contents
#     driver.find_element(By.CLASS_NAME, 'dep1').click()
#     time.sleep(0.5)
#     start_time_Entity03 = time.time()
#     driver.find_element(By.CLASS_NAME, 'x-grid-tree-node-leaf').click()
#     end_time_Entity03 = time.time()
#     entity03 = end_time_Entity03 - start_time_Entity03
#     print("Main Page>All Contents Page : " + str(entity03))
#     teams_section.addFact("Main Page>All Contents Page", str(entity03))

#     ## My Workplace 이동
#     start_time_Entity04 = time.time()
#     driver.find_element(By.XPATH, '//*[@id="ext-element-1"]/div[1]/div[2]/div[1]/ul[2]/li[1]/a').click()
#     end_time_Entity04 = time.time()
#     entity04 = end_time_Entity04 - start_time_Entity04

#     print("All Contents Page>My Workplace Page : " + str(entity04))
#     teams_section.addFact("All Contents Page>My Workplace Page", str(entity04))    

#     teamsMsg.title("페이지 호출 테스트")
#     teamsMsg.text("실행 시간 : " + currentTime)
#     teamsMsg.addSection(teams_section)
#     teamsMsg.send()

#     print("(Page)테스트 종료")
#     time.sleep(1)
#     # fileUploadTest.exec()


# # 매일 08시 30분에 함수 실행
# # schedule.every().day.at("15:10").do(main)

# # while True:
# #     schedule.run_pending()
    
# main()
