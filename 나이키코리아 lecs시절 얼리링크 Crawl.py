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
 

 

# xlwt(�������� pip)�� �̿��� ���ο� ������ workbook�� utf-8�� ���ڵ��Ͽ� ����.

# ������ workbook > worksheet > cell ������ ����!

wb = xlwt.Workbook(encoding='utf-8')

# ���Ӱ� �� ���� wb�� worksheet �̸��� '2017_10 early link'�� ����

sheet = wb.add_sheet('2017_10 early link')

 

 

#����̹��� ũ��������̹��� �̿�

#������������� ũ�Ѹ��� �ϰ��� �ϱ⿡ ����������� ���� useragent�� ����ϱ��� ����

mobile_emulation = {

    "deviceMetrics": { "width": 360, "height": 640, "pixelRatio": 3.0 },

    "userAgent": "Mozilla/5.0 (Linux; Android 4.2.1; en-us; Nexus 5 Build/JOP40D) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.166 Mobile Safari/535.19" }

chrome_options = Options()

chrome_options.add_argument("--disable-popup-blocking");

chrome_options.add_argument("test-type");

chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)

driver = webdriver.Chrome(chrome_options = chrome_options)

 

 

#for�� �ȿ��� ����� j=0�� for�� �ȿ� �������ָ�

#�ݺ����� ���� ��� 0���� �ʱ�ȭ�Ǳ⿡ �ۿ� ��������

j=0

 

#for �ݺ����� �����µ� startnum�̶� ������ ������ 31090330~31099999�� �����Ͽ�

#startnum������ 31090330������ 31090331, 31090332�� ���� ������ �ش� ������ �������� ����.

#startnum = ��ǰ�ѹ��μ� URL�� �ִ� ��ǰ�ѹ��� ���� �������� �ش� ��ǰ�������� �޶�����.

for startnum in range(31100000, 31109999):

        #ũ�Ѹ� �� URL �� NK���� ���� startnum�� ���� ������ ���� ���� 

        #%d�� ���� %d�� ���� startnum���� ����

        driver.get("http://m.nike.co.kr/mobile/goods/showGoodsDetail.lecs?goodsNo=NK%d" % startnum)

        #Try except���� ����

        #ũ��������̹��� alert�� �Ͼ ��� alert�� accept�ϰ�

        #alert�� ���� ��� pass��

        #except���� ���� alert.accept()�� �ۼ��� ��� ������ �߻����� ���� ��찡 �����ΰ�ó�� ���� ����ȴ�.

        try:

            alert = driver.switch_to_alert()

            alert.accept()

        except:

            pass

        #ũ�Ѹ��ϰ��� �ϴ� URL(driver.get�� ������ �ּ�) �� crawl�ϰ��� �ϴ� �κ��� �����ڵ����� ���� ã�� ��

        #xpath��θ� �������־� xpath�� ���� �ش� �κ��� ã���� �Ѵ�.        

        page_results = driver.find_elements_by_xpath('//*[@id="productDetail"]/div[1]/h1')

        page_results

        

        #page_results ����� �� Ŭ���� = tit (���� xpath�ؼ� ã�����ϴ� �κ�)����

        for tit in page_results:

            print(tit.text) #Ŭ���� = tit�� ���� text�� print��

            sheet.write(j,1,tit.text) #�ش� text�� ���� �� worksheet�� 0,1(B1��)�� write�Ѵ�.

            j+=1 #j�� �ϳ��� ������ 1,1(B2��), 2,1(B3��)�� ���� tit.text���� �������� ���� �������� write�Ҽ��ֵ��� �Ѵ�.

            wb.save('2017_11_Early_Link_Crawl_After_31100000_5.xls') #"2017_09_Early_Link_Crawl.xls"�̸����� �ش� ������ ���̺��Ѵ�.