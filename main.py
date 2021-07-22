from selenium import webdriver
import time
from bs4 import BeautifulSoup
import openpyxl
import os, sys
import util
from selenium.webdriver.support.ui import Select
import crawling, crawling2



print("엑셀 파일 가져오는 중")
all_values=util.excel()


options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
# options.add_argument('headless')
# options.add_argument("no-sandbox")
driver = webdriver.Chrome(options=options)

#crawling.hana(all_values,driver)
#crawling.haeorem(all_values,driver)

crawling2.vkfclf(all_values,driver)


print("########################## 프로그램 완료 ##################################")
