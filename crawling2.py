from selenium import webdriver
import time
from bs4 import BeautifulSoup
import openpyxl
import os, sys
import util
from selenium.webdriver.support.ui import Select




exceptLine = []
reason = []

def vkfclf(all_values,driver):
    url = 'http://www.8725.com/new/confirm_login.php'
    driver.get(url)
    driver.implicitly_wait(3)

    driver.find_element_by_name('id').send_keys('jesus1018')
    driver.find_element_by_name('passwd').send_keys('a89838983!')
    driver.find_element_by_xpath('/html/body/div/form/button').click()
    driver.implicitly_wait(3)
    driver.get('http://www.8725.com/8725_admin/supply_product/supply_product_new.php')
    excep = []

    for i in range(2,len(all_values)):
        Select(driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td/form/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/select')).select_by_visible_text(text='전체상품')
        time.sleep(0.5)
        Select(driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td/form/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/select')).select_by_visible_text(text='검토중상품')
        time.sleep(0.5)
        Select(driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td/form/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/select')).select_by_visible_text(text='상품명')
        time.sleep(0.5)
        driver.find_element_by_name('searchstring').send_keys(all_values[i][13])
        driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td/form/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/input').click()
        if len(driver.find_elements_by_xpath('/html/body/table/tbody/tr[2]/td/form/table/tbody/tr/td[3]/table/tbody/tr[4]/td/table/tbody/tr[3]/td/table/tbody/tr[*]/td[4]'))>1:
            excep.append(i)
            print("중복된 상품을 찾았습니다.")







    for i in range(2,len(all_values)):
        if not excep:
            print("")
        else :
            if i == excep.pop(0):
                print(all_values[i][13]+" 이미등록하였습니다.########")
                continue


        driver.implicitly_wait(3)
        driver.get('http://www.8725.com/8725_admin/supply_product/supply_product_write.php')



        driver.find_element_by_name('name1').send_keys(all_values[i][13])
        driver.find_element_by_name('r_sale').send_keys(all_values[i][18])

        driver.find_element_by_name('quantity').send_keys(all_values[i][26])
        driver.find_element_by_name('printer').send_keys(all_values[i][22])



        driver.find_element_by_name('wrap').send_keys(all_values[i][29])
        driver.find_element_by_name('origin').send_keys('중국')
        Select(driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td/form/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td/table/tbody/tr[3]/td[4]/span/select')).select_by_visible_text(text='10 일')
        driver.find_element_by_name('print_won').send_keys('30000')
        driver.find_element_by_name('s_pace').send_keys(all_values[i][21])
        driver.find_element_by_name('wood').send_keys(all_values[i][28])
        driver.find_element_by_name('color').send_keys(all_values[i][14])
        driver.find_element_by_name('size').send_keys(all_values[i][30])
        driver.find_element_by_name('h_stock').send_keys(all_values[i][27])
        driver.find_element_by_name('h_money').send_keys('4000')


        try:
            if '실크1도인쇄' in str(all_values[i][22]):
                driver.find_element_by_name('addoption1').send_keys('실크인쇄')
            elif '스티커' in str(all_values[i][22]):
                driver.find_element_by_name('addoption1').send_keys('스티커')
            elif '레이저인쇄' in str(all_values[i][22]):
                driver.find_element_by_name('addoption1').send_keys('레이저인쇄')


            driver.find_element_by_name('add1').send_keys('300')
            if '종이케이스' in str(all_values[i][29]):
                driver.find_element_by_name('addoption2').send_keys('선물포장지')
            driver.find_element_by_name('add2').send_keys('300')
        except:
            exceptLine.append(int(i)+1)
            reason.append('인쇄 및 케이스 확인')
            url = 'http://cpsite.jclgift.com/'
            driver.get(url)
            time.sleep(1)
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 인쇄 및 케이스 확인 #!#!#!')
            continue


        try:
            colores = []
            colores = str(all_values[i][14]).split(',')
            count=3
            for col in colores:
                driver.find_element_by_name('addoption'+str(count)).send_keys(col.replace(" ",""))
                count=count+1
        except:
            exceptLine.append(int(i)+1)
            reason.append('컬러 확인')
            url = 'http://cpsite.jclgift.com/'
            driver.get(url)
            time.sleep(1)
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 컬러 확인 #!#!#!')
            continue


        driver.find_element_by_name('etc').send_keys('대량구매환영')
        driver.find_element_by_name('admin').send_keys('인쇄 : 1도 인쇄\n인쇄비 : 개당 300원\n100개 이하 인쇄비 : 30,000원\n1박스당 택배비 : 4,000원\n공급처 상호 : 3D토탈기획\n공급처 전화 : 010-6444-8983\n공급처 메일 : 3d898@naver.com')

        try:
            fullName=util.findImage(all_values[i][3],'135101')
            driver.find_element_by_name('userfile0').send_keys(fullName)
            fullName=util.findImage(all_values[i][3],'800600')
            driver.find_element_by_name('userfile3').send_keys(fullName)
            fullName=util.findImage(all_values[i][3],'detail')
            driver.find_element_by_name('userfile4').send_keys(fullName)
            fullName=util.findImage(all_values[i][3],'sample')
            driver.find_element_by_name('userfile5').send_keys(fullName)
            fullName=util.findImage(all_values[i][3],'500500')
            driver.find_element_by_name('userfile8').send_keys(fullName)
        except:
            exceptLine.append(int(i)+1)
            reason.append('이미지 이름 및 경로 확인')
            url = 'http://cpsite.jclgift.com/'
            driver.get(url)
            time.sleep(1)
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 이미지 이름 및 경로 확인 #!#!#!')
            continue

        time.sleep(1)
        driver.execute_script('javascript:list_edit(this.form);')
        driver.implicitly_wait(3)
        driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td/form/table/tbody/tr/td[3]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[13]/td/table/tbody/tr/td[1]/input').click()
        time.sleep(1)
        print("####### "+str(all_values[i][0])+' / '+ str(len(all_values)-2) +' 완료 #####')
