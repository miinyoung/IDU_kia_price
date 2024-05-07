# #-*- coding: utf-8 -*-
# ### 2024.02.16 
# ### 작성자 : 유정원

# pip install selenium
# pip install webdriver-manager
# pip install configparser
# pip install openpyxl

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from openpyxl import Workbook
from datetime import datetime
import os
import configparser
import time
import copy

def pricePrint(row):
    paymentPrice = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div/div[4]/h3/button/span[2]/span[1]').text
    paymentPrice = paymentPrice[:-1]
    paymentPrice = paymentPrice.replace(",", "")
    row.append(paymentPrice)    
    print("결제 금액 : " + paymentPrice)

    taxPrice = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div/div[5]/h3/button/span[2]/span[1]').text
    taxPrice = taxPrice[:-1]
    taxPrice = taxPrice.replace(",", "")
    row.append(taxPrice)
    print("면세 구분 및 등록비 : " + taxPrice)
    acquisitionPrice = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div/div[5]/div/div[2]/div/div[2]').text
    acquisitionPrice = acquisitionPrice[:-1]
    acquisitionPrice = acquisitionPrice.replace(",", "")
    row.append(acquisitionPrice)
    print("취득세 : " + acquisitionPrice)
    bondPrice = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div/div[5]/div/div[3]/div[1]/div[2]').text
    bondPrice = bondPrice[:-1]
    bondPrice = bondPrice.replace(",", "")
    row.append(bondPrice)
    print("공채 : " + bondPrice)

    licensePrice = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div/div[5]/div/div[3]/div[3]/div[1]/div[2]').text
    licensePrice = licensePrice[:-1]
    licensePrice = licensePrice.replace(",", "")
    row.append(licensePrice)
    print("차량 번호판 : " + licensePrice)

    registPrice = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div/div[5]/div/div[3]/div[3]/div[2]/div[2]').text
    registPrice = registPrice[:-1]
    registPrice = registPrice.replace(",", "")
    row.append(registPrice)
    print("등록 대행 수수료 : " + registPrice)

    try:
        mTaxPrice = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div/div[5]/div/div[4]/div[1]/div[2]').text
        mTaxPrice = mTaxPrice[1:-1]#- 부호 생략
        mTaxPrice = mTaxPrice.replace(",", "")
        mTaxPrice = mTaxPrice.replace(" ", "")
        row.append(mTaxPrice)
        print("세금 변동 : " + mTaxPrice)
    except:
        row.append("0") #없는 경우 0
        pass

    print(row)
    resultRows.append(row)


def main(car):
    row = []
    trimSelected = ""
    carName = car["NAME"]
    waitS = 5 #기다리는 시간
    row.append(carName)
    print(carName)

    # 크롤러 작동부
    URL = car["URL"]
    driver.get(URL)

    driver.find_element(By.CLASS_NAME, 'conf-btn-close').click()
    driver.implicitly_wait(3)

    ## 인승 선택
    peopleNum = str(car["NUM"])
    if peopleNum == "none":
        pass     
    else : 
        number = str("//label[@for='" + car["NUM"] + "']")
        driver.find_element(By.XPATH, number).click()
        trimSelected += driver.find_element(By.XPATH, number).text + "/"
        driver.implicitly_wait(3)

    ## 엔진 선택
    engine = str(car["ENGINE"])
    if engine == 'none':
        pass
    else:
        engine = str("//label[@for='" + car["ENGINE"] + "']")
        driver.find_element(By.XPATH, engine).click()
        trimSelected += driver.find_element(By.XPATH, engine).text + " / "
        driver.implicitly_wait(3)    

    ## 미션 선택
    trans = str(car["TRANSMISSION"])
    if trans == 'none':
        pass
    else:
        trans = str("//label[@for='" + car["TRANSMISSION"] + "']")
        driver.find_element(By.XPATH, trans).click()
        trimSelected += driver.find_element(By.XPATH, trans).text + " / "
        driver.implicitly_wait(3)    

    ## 트림
    trim = str("//label[@for='" + car["TRIM"] + "']")
    driver.find_element(By.XPATH, trim).click()
    trimSelected += driver.find_element(By.XPATH, trim+"/div/strong").text + " / "
    driver.implicitly_wait(3)    

    ## 컬러 선택 이동
    driver.find_element(By.XPATH, '//*[@id="build_menu_btns"]/div/div/div/div/button').click()
    driver.implicitly_wait(3)    

    ## 외장 컬러 선택
    extColor = str("//label[@for='" + car["EXT_COLOR"] + "']")
    driver.find_element(By.XPATH, extColor).click()
    trimSelected += driver.find_element(By.XPATH, '//*[@id="configurator_menu_content"]/div[1]/div/div/div/div[4]/div[1]/div[1]').text + " / "
    driver.implicitly_wait(3)    

    ## 내장 컬러 선택
    intColor = str("//label[@for='" + car["INT_COLOR"] + "']")
    driver.find_element(By.XPATH, intColor).click()
    trimSelected += driver.find_element(By.XPATH, '//*[@id="configurator_menu_content"]/div[1]/div/div/div/div[6]/div[1]/div').text + " / "    
    driver.implicitly_wait(3)    

    ## 옵션 선택
    driver.find_element(By.XPATH, '//*[@id="build_menu_btns"]/div/div/div/div[2]/button').click()
    driver.implicitly_wait(3)

    ## 패키지 선택
    packageNum = str(car["PACKAGE"])
    if packageNum == "none":
        pass
    else:
        package = str("//label[@for='" + car["PACKAGE"] + "']")
        driver.find_element(By.XPATH, package).click()
        trimSelected += driver.find_element(By.XPATH, package).text + " / "
        driver.implicitly_wait(3)        

    ## 상세 옵션 선택
    if str(car["OPTION"]) == "none":
        pass
    else :
        option = str("//label[@for='" + car["OPTION"] + "']")
        driver.find_element(By.XPATH, option).click()
        trimSelected += driver.find_element(By.XPATH, option + "/span[2]").text
        driver.implicitly_wait(3)

    ## 선택 내용 추가
    row.append(trimSelected)   

    ## 액세사리 선택(없음)
    driver.find_element(By.XPATH, '//*[@id="build_menu_btns"]/div/div/div/div[2]/button').click()
    driver.implicitly_wait(3)

    ## 견적 완료
    driver.find_element(By.XPATH, '//*[@id="build_menu_btns"]/div/div/div/div[2]/button').click()
    time.sleep(RESULTPAGE_LOADTIME) #계산 때문에 3초 딜레이 추가

    ## 면세 구분 클릭
    button_element = driver.find_element(By.XPATH, '//button[@data-link-label="면세 구분 및 등록비_상세보기"]')
    driver.execute_script("arguments[0].click();", button_element)
    time.sleep(RESULTPAGE_LOADTIME)

    ### 일반인
    row.append("일반인")
    pricePrint(row.copy())

    ### 일반인 + 다자녀
    row[2] = "일반인+다자녀"
    checkBox = driver.find_element(By.XPATH, "//label[@for='multiChildYn']")
    driver.execute_script("arguments[0].click();", checkBox)
    time.sleep(RESULTPAGE_LOADTIME)
    pricePrint(row.copy())

    ## 면세선택
    dropdown_button = driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div[2]/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div/div[5]/div/div[1]/div[2]/div[1]")
    driver.execute_script("arguments[0].click();", dropdown_button)
    time.sleep(RESULTPAGE_LOADTIME)

    ### 장애 4~6 등급
    row[2] = "장애 4~6 등급"
    list01 = driver.find_element(By.XPATH, "//li[@class='dropdown-select__item']/button[text()='장애 4~6급']")
    driver.execute_script("arguments[0].click();", list01)
    time.sleep(RESULTPAGE_LOADTIME)
    pricePrint(row.copy())

    ## 국가유공자
    row[2] = "국가유공자"
    list02 = driver.find_element(By.XPATH, "//li[@class='dropdown-select__item']/button[text()='국가유공자']")
    driver.execute_script("arguments[0].click();", list02)
    time.sleep(RESULTPAGE_LOADTIME)
    pricePrint(row.copy())

    ## 광주민주화
    row[2] = "광주민주화"
    list03 = driver.find_element(By.XPATH, "//li[@class='dropdown-select__item']/button[text()='광주민주화']")
    driver.execute_script("arguments[0].click();", list03)
    time.sleep(RESULTPAGE_LOADTIME)
    pricePrint(row.copy())

    # for eachRow in resultRows:
    #     resultFile.append(eachRow)

    # wb.save(os.path.join(os.getcwd(), "result/") + "KiaPrice_" + currentTime+ ".xlsx")
    # print("---------------------------------------------")
    # time.sleep(0.5)
    
## Excel setting
wb = Workbook()
resultFile = wb.active
resultRows = []
# currentTime = str(datetime.today().strftime("%Y%m%d_%H%M")) 
currentTime = str(datetime.today().strftime("%Y-%m-%d")) 

resultFile["A1"].value = "차량"
resultFile["B1"].value = "선택 내용"
resultFile["C1"].value = "면세 구분"
resultFile["D1"].value = "결제 금액"
resultFile["E1"].value = "면세 구분 및 등록비"
resultFile["F1"].value = "취득세"
resultFile["G1"].value = "공채"
resultFile["H1"].value = "차량 번호판"
resultFile["I1"].value = "등록 대행 수수료"
resultFile["J1"].value = "세금 변동"

## Read setting.ini
properties = configparser.ConfigParser()
print(properties)
properties.read('setting.ini') # 경로
print(properties.sections())

default = properties["Default"]
Sportage = properties["Sportage"]
K8 = properties["K8"]
Sorento = properties["Sorento"]
EV6 = properties["EV6"]
EV9 = properties["EV9"]
Carnival = properties["Carnival"]

global RESULTPAGE_LOADTIME
RESULTPAGE_LOADTIME = int(default["RESULTPAGE_LOADTIME"])

carList = [Sportage, K8, Sorento, EV6, EV9, Carnival]

## selenuim Chrome Driver Option Setting
options = webdriver.ChromeOptions()
options.add_argument('lang=ko_KR') # 사용언어 한국어
# options.add_argument('disable-gpu') # 하드웨어 가속 안함
# options.add_argument('log-level=3')
options.add_argument('--window-size=1920,1080') # 해상도 설정 
options.add_argument('headless')
# options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
options.add_experimental_option("detach",True) # 웹브라우저 유지

service = Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

for car in carList:
    main(car)

for eachRow in resultRows:
    resultFile.append(eachRow)


wb.save(os.path.join(os.getcwd(), "result/") + "KiaPrice_" + currentTime+ ".xlsx") #경로 
print("---------------------------------------------")
time.sleep(0.5)

print("------------End------------")