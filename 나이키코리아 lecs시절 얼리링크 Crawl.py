# This Python file uses the following encoding: utf-8
 

import os
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert
from openpyxl import workbook
import xlrd
import xlwt
from xlutils.copy import copy
from selenium.common.exceptions import NoSuchElementException,StaleElementReferenceException
 

 

# xlwt(엑셀관련 pip)를 이용해 새로운 엑셀의 workbook을 utf-8로 인코딩하여 연다.

# 엑셀은 workbook > worksheet > cell 순서로 접근!

wb = xlwt.Workbook(encoding='utf-8')

# 새롭게 연 엑셀 wb에 worksheet 이름을 '2017_10 early link'로 만듦

sheet = wb.add_sheet('2017_10 early link')

 

 

#드라이버는 크롬웹드라이버를 이용

#모바일페이지로 크롤링을 하고자 하기에 모바일접근을 위해 useragent를 모바일기기로 설정

mobile_emulation = {

    "deviceMetrics": { "width": 360, "height": 640, "pixelRatio": 3.0 },

    "userAgent": "Mozilla/5.0 (Linux; Android 4.2.1; en-us; Nexus 5 Build/JOP40D) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.166 Mobile Safari/535.19" }

chrome_options = Options()

chrome_options.add_argument("--disable-popup-blocking");

chrome_options.add_argument("test-type");

chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)

driver = webdriver.Chrome(chrome_options = chrome_options)

 

 

#for문 안에서 사용할 j=0을 for문 안에 정의해주면

#반복문을 돌며 계속 0으로 초기화되기에 밖에 선언해줌

j=0

 

#for 반복문을 돌리는데 startnumber부터 finishnumber까지

#startnum변수부터 1씩 오르며 해당 범위를 돈다.

#startnum = 상품넘버로서 URL에 있는 상품넘버에 따라 페이지가 해당 상품페이지로 달라진다.

for startnum in range(startnumber, finishnumber): # need to set range

        #크롤링 할 URL 중 NK뒤의 값이 startnum이 가진 변수로 돌기 위해 

        #%d로 놓고 %d의 값이 startnum임을 선언

        driver.get("http://m.nike.co.kr/mobile/goods/showGoodsDetail.lecs?goodsNo=NK%d" % startnum)

        #Try except문의 설명

        #크롬웹드라이버에 alert이 일어날 경우 alert을 accept하고

        #alert이 없을 경우 pass함

        #except문이 없이 alert.accept()만 작성할 경우 에러가 발생하지 않은 경우가 에러인것처럼 실행 종료된다.

        try:

            alert = driver.switch_to_alert()

            alert.accept()

        except:

            pass

        #크롤링하고자 하는 URL(driver.get에 정의한 주소) 내 crawl하고자 하는 부분을 개발자도구를 통해 찾은 후

        #xpath경로를 지정해주어 xpath를 통해 해당 부분을 찾도록 한다.        

        page_results = driver.find_elements_by_xpath('//*[@id="productDetail"]/div[1]/h1')

        page_results

        

        #page_results 결과값 내 클래스 = tit (제가 xpath해서 찾고자하는 부분)에서

        for tit in page_results:

            print(tit.text) #클래스 = tit인 값의 text를 print함

            sheet.write(j,1,tit.text) #해당 text를 엑셀 내 worksheet의 0,1(B1셀)에 write한다.

            j+=1 #j를 하나씩 높여가 1,1(B2셀), 2,1(B3셀)과 같이 tit.text값을 차근차근 셀을 내려가며 write할수있도록 한다.

            wb.save('2017_11_Early_Link_Crawl_After_31100000_5.xls') #"2017_09_Early_Link_Crawl.xls"이름으로 해당 파일을 세이브한다.
