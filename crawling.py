from selenium import webdriver
import time
from bs4 import BeautifulSoup
import openpyxl
import os, sys
import util
from selenium.webdriver.support.ui import Select




exceptLine = []
reason = []

def hana(all_values,driver):
    url = 'https://gdadmin.bizpanchok.co.kr/base/login.php'
    driver.get(url)
    driver.implicitly_wait(3)
    driver.find_element_by_name('managerId').send_keys('2220859859')
    driver.find_element_by_name('managerPw').send_keys('a89838983!')
    driver.find_element_by_xpath('//*[@id="frmLogin"]/table/tbody/tr/td/div/div[1]/div[2]/input').click()
    time.sleep(3)
    driver.implicitly_wait(3)


    count = (len(all_values) - 2)
    for i in range(2,len(all_values)):
        driver.implicitly_wait(5)
        driver.get('https://gdadmin.bizpanchok.co.kr/provider/goods/goods_register.php')

        time.sleep(1)
        sec=0.5


        try:

            flag = 0
            for ti in driver.find_elements_by_xpath('//*[@id="cateGoods1"]/*'):
                if str(all_values[i][47]).replace(" ","") in ti.text:
                    Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                    flag = 1
                    break
            time.sleep(sec)
            if flag == 0:
                raise exception()
            flag = 0
            for ti in driver.find_elements_by_xpath('//*[@id="cateGoods2"]/*'):
                if str(all_values[i][48]).replace(" ","") in ti.text:
                    Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                    flag = 1
                    break
            time.sleep(sec)
            if flag == 0:
                raise exception()
            flag = 0
            for ti in driver.find_elements_by_xpath('//*[@id="cateGoods3"]/*'):
                if str(all_values[i][49]).replace(" ","") in ti.text:
                    Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                    flag = 1
                    break
            time.sleep(sec)
            if flag == 0:
                raise exception()
            flag = 0
            for ti in driver.find_elements_by_xpath('//*[@id="cateGoods4"]/*'):
                if str(all_values[i][50]).replace(" ","") in ti.text:
                    Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                    flag = 1
                    break
            time.sleep(sec)
            if flag == 0:
                raise exception()
        except:

            exceptLine.append(int(i)+1)
            reason.append('카테고리 확인')
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 카테고리 확인 #!#!#!')
            time.sleep(1)
            continue



        try:
            driver.find_element_by_xpath('//*[@id="btn_category_select"]').click()
            time.sleep(sec)
            driver.find_elements_by_name('cateCd')[3].click()


            driver.find_element_by_name('goodsNm').send_keys(all_values[i][13])
            driver.find_element_by_name('addGoodsKeyword').click()
            driver.find_element_by_name('originNm').send_keys('중국')
            driver.find_element_by_name('sampleFl').click()
            time.sleep(sec)
            driver.find_element_by_name('sampleSupplyPrice').send_keys(all_values[i][18])
            driver.find_element_by_xpath('//*[@id="sampleGroupTitleInfo"]/tbody/tr[2]/td/label/button').click()

            driver.find_element_by_name('deliverChk').click()
            driver.find_element_by_xpath('//*[@id="layer-wrap"]/div[4]/input').click()
            driver.find_element_by_name('corp1Fl').click()

            driver.find_element_by_name('unit1Cnt').clear()
            driver.find_element_by_name('unit1Cnt').send_keys(all_values[i][24])
            driver.find_element_by_name('unit2Cnt').send_keys(all_values[i][25])
            driver.find_element_by_name('unit3Cnt').send_keys(all_values[i][26])
            driver.find_element_by_name('unit4Cnt').send_keys(all_values[i][27])
            driver.find_element_by_name('unit5Cnt').send_keys(all_values[i][28])

            driver.find_element_by_name('unit1SupplyPrice').send_keys(all_values[i][18])
            driver.find_element_by_name('unit2SupplyPrice').send_keys(all_values[i][18])
            driver.find_element_by_name('unit3SupplyPrice').send_keys(all_values[i][18])
            driver.find_element_by_name('unit4SupplyPrice').send_keys(all_values[i][18])
            driver.find_element_by_name('unit5SupplyPrice').send_keys(all_values[i][18])
        except:
            exceptLine.append(int(i)+1)
            reason.append('웹사이트 확인')
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 웹사이트 확인 #!#!#!')
            time.sleep(1)
            continue


        try:
            if '포장없음' in str(all_values[i][29]):
                driver.find_elements_by_name('optionPackingFl')[1].click()
            elif '자체케이스' in str(all_values[i][29]):
                driver.find_elements_by_name('optionPackingFl')[1].click()



            else:
                flag = 0
                for ti in driver.find_elements_by_xpath('//*[@id="devOptionPacking1"]/*'):
                    if str(all_values[i][29]).replace(" ","") in ti.text:
                        Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                        flag = 1
                        break
                Select(ti.find_element_by_xpath('//*[@id="optionPackingForm1"]/td[3]/div/select')).select_by_visible_text(text='노출함')

                if '종이케이스' in str(all_values[i][29]):
                    driver.find_element_by_xpath('//*[@id="optionPackingForm1"]/td[1]/div/input[2]').click()
                    time.sleep(sec)
                    Select(ti.find_element_by_xpath('//*[@id="devOptionPacking2"]')).select_by_visible_text(text='선물포장지포장')
                    driver.find_elements_by_name('optionPacking[unitPrice][]')[1].clear()
                    driver.find_elements_by_name('optionPacking[unitPrice][]')[1].send_keys('300')
                    Select(ti.find_element_by_xpath('//*[@id="optionPackingForm2"]/td[3]/div/select')).select_by_visible_text(text='노출함')
        except:
            exceptLine.append(int(i)+1)
            reason.append('케이스 확인')
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 케이스 확인 #!#!#!')
            time.sleep(1)
            continue



        try:
            driver.find_elements_by_name('optionStickerFl')[1].click()
            if '스티커' in str(all_values[i][22]):
                driver.find_elements_by_name('optionPrintFl')[1].click()
            else :
                if '실크인쇄' in str(all_values[i][22]):
                    Select(ti.find_element_by_xpath('//*[@id="devOptionPrint1"]')).select_by_visible_text(text='실크1도 인쇄')
                if '레이저인쇄' in str(all_values[i][22]):
                    Select(ti.find_element_by_xpath('//*[@id="devOptionPrint1"]')).select_by_visible_text(text='레이저인쇄')


                driver.find_element_by_name('optionPrint[underCnt][]').clear()
                driver.find_element_by_name('optionPrint[underCnt][]').send_keys('100')
                driver.find_element_by_name('optionPrint[bundlePrice][]').clear()
                driver.find_element_by_name('optionPrint[bundlePrice][]').send_keys('30000')
                driver.find_element_by_name('optionPrint[unitPrice][]').clear()
                driver.find_element_by_name('optionPrint[unitPrice][]').send_keys('300')
                Select(ti.find_element_by_xpath('//*[@id="optionPrintForm1"]/td[4]/div/select')).select_by_visible_text(text='노출함')



        except:
            exceptLine.append(int(i)+1)
            reason.append('스티커 및 인쇄 확인')
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 스티커 및 인쇄 확인 #!#!#!')
            time.sleep(1)
            continue

        try:

            fullName =util.findImage(str(all_values[i][3]))
            #driver.find_element_by_name('imageResize[original]').click()
            driver.find_element_by_xpath('//*[@id="filestyle-0"]').send_keys(fullName)

        except:
            exceptLine.append(int(i)+1)
            reason.append('이미지 이름 및 경로 확인')
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 이미지 이름 및 경로 확인 #!#!#!')
            time.sleep(1)
            continue

        try:
            driver.switch_to.frame(driver.find_element_by_xpath('//*[@id="textareaDescriptionShop"]/iframe'))
            time.sleep(sec)
            driver.find_element_by_xpath('//*[@id="smart_editor2_content"]/div[4]/ul/li[2]/button').click()
            #driver.switch_to.frame(driver.find_element_by_xpath('//*[@id="se2_iframe"]'))
            driver.find_element_by_xpath('//*[@id="smart_editor2_content"]/div[3]/textarea[1]').send_keys(all_values[i][31])
            driver.find_element_by_xpath('//*[@id="smart_editor2_content"]/div[3]/textarea[1]').send_keys('\n\n')
            driver.find_element_by_xpath('//*[@id="smart_editor2_content"]/div[3]/textarea[1]').send_keys(all_values[i][32])
            driver.switch_to_default_content()


            driver.find_element_by_name('externalVideoFl').click()
            time.sleep(sec)

            driver.find_element_by_name('externalVideoUrl').send_keys(all_values[i][36])

            driver.find_element_by_name('memo').send_keys('제작기간 10~11일\n인쇄 : 1도 인쇄\n인쇄비 : 개당 300원\n100개 이하 인쇄비 : 30,000원\n1박스당 택배비 : 4,000원\n공급처 상호 : 3D토탈기획\n공급처 전화 : 010-6444-8983\n공급처 메일 : 3d898@naver.com')


            driver.find_element_by_name('frmGoods').submit()

            time.sleep(sec*7)


        except:
            exceptLine.append(int(i)+1)
            reason.append('웹사이## 확인')
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 이미지 이름 및 경로 확인 #!#!#!')
            time.sleep(1)
            continue
        print("####### "+str(all_values[i][0])+' / '+ str(count) +' 완료 #####')

def haeorem(all_values,driver):


    url = 'http://cpsite.jclgift.com/'
    driver.get(url)
    time.sleep(1)


    try:
        result= driver.switch_to.alert

        result.dismiss()
        result.aceept()


    except:
        print("    ")


    #driver.implicitly_wait(time_to_wait=5)

    driver.implicitly_wait(3)

    driver.find_element_by_name('Mid').send_keys('jesus1018')
    driver.find_element_by_name('Pwd').send_keys('0000')
    driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr/td[3]/p/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td[3]/input').click()

    count = (len(all_values) - 2)
    for i in range(2,len(all_values)):
        ###등록 실패
        try:
            driver.implicitly_wait(5)

            driver.switch_to_window(driver.window_handles[0])
            driver.switch_to.frame(driver.find_element_by_name('jobFrame'))
            #soup = BeautifulSoup(driver.page_source, "html.parser")
            #print(soup)

            #해당 script 바로 실행
            driver.execute_script("location.href='product_add.asp'")
        except:

            reason.append('사이트가 불안정합니다.')
            exceptLine.append(int(i)+1)
            url = 'http://cpsite.jclgift.com/'
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 사이트가 불안정합니다 #!#!#!')

            driver.get(url)
            continue
        ###분류 실패
        try:
            #창 닫기
            driver.switch_to_window(driver.window_handles[1])
            driver.close()
            driver.switch_to_window(driver.window_handles[0])
            driver.switch_to.frame(driver.find_element_by_name('jobFrame'))

            driver.find_element_by_name('p_name').send_keys(all_values[i][13])
            driver.find_element_by_name('p_modelname').send_keys(all_values[i][13])
            driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[4]/td[2]/table/tbody/tr[1]/td[2]/img').click()
            time.sleep(1)

            try:
                result= driver.switch_to.alert

                result.dismiss()
                result.aceept()


            except:
                print(" ")
            #상품
            driver.find_element_by_name('p_gusung').send_keys('상세페이지 참조')
            driver.find_element_by_name('p_size').send_keys(all_values[i][30])
            driver.find_element_by_name('p_material').send_keys('상세페이지 참조')
            driver.find_element_by_name('p_color').send_keys(all_values[i][14])

            flag = 0
            #Select(driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select[1]')).select_by_visible_text(text=all_values[i][45])
            for ti in driver.find_elements_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select[1]/*'):
                if all_values[i][45] in ti.text:
                    Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                    flag = 1
                    break
            if flag == 0:
                raise exception()
            flag = 0
            for ti in driver.find_elements_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select[2]/*'):
                if all_values[i][46] in ti.text:
                    Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                    flag = 1
                    break
            if flag == 0:
                raise exception()
            flag = 0
            for ti in driver.find_elements_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select[3]/*'):
                if all_values[i][47] in ti.text:
                    Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                    flag = 1
                    break

            if flag == 0:
                raise exception()
            flag = 0
            for ti in driver.find_elements_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[3]/td[2]/select[1]/*'):
                if all_values[i][48] in ti.text:
                    Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                    flag = 1
                    break
            if flag == 0:
                raise exception()
            flag = 0
            for ti in driver.find_elements_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[3]/td[2]/select[2]/*'):
                if all_values[i][49] in ti.text:
                    Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                    flag = 1
                    break
            if flag == 0:
                raise exception()
            flag = 0
            for ti in driver.find_elements_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[3]/td[2]/select[3]/*'):
                if all_values[i][50] in ti.text:
                    Select(ti.find_element_by_xpath('..')).select_by_visible_text(text=ti.text)
                    flag = 1
                    break
            if flag == 0:
                raise exception()
            flag = 0




            driver.find_element_by_name('p_origin').send_keys('중국')
            driver.find_element_by_name('p_allow').send_keys('해당없음')
            driver.find_element_by_name('p_makeperiod').clear()
            driver.find_element_by_name('p_makeperiod').send_keys('디자인 시안확정 후 10~11일')
            driver.find_element_by_name('p_box').send_keys('별도문의')
            driver.find_element_by_name('quantity_set').send_keys(all_values[i][26])
            driver.find_element_by_name('p_price1').send_keys(all_values[i][18])
            driver.find_element_by_name('opt_free_print_count').send_keys('0')
            driver.find_element_by_name('p_company_fax').send_keys('0303-3130-8983')


            #driver.find_element_by_name('p_company_memo').send_keys('기본수량 : '+str(all_values[i][26])+'개 \n'+'ddd')
            driver.find_element_by_name('p_company_memo').send_keys('인쇄 : 1도 인쇄 \n인쇄비 : 개당 300원\n100개 이하 인쇄비 : 30,000원\n1박스당 택비비 : 4,000원')


            if '스티커' in str(all_values[i][22]):
                driver.find_element_by_name('opt_use_YN').click()
                flag=1
            else :
                for ti in driver.find_elements_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[9]/td[2]/table/tbody/tr[*]/td[1]'):
                    if str(all_values[i][22]) in ti.text:
                        ti.find_element_by_xpath('./input').click()
                        #time.sleep(1)
                        ti.find_element_by_xpath('../td[3]/input').send_keys(int(all_values[i][23]))
                        flag = 1
                        break
                if flag == 0:
                    raise exception()
            flag = 0

        except:


            exceptLine.append(int(i)+1)
            reason.append('분류, 인쇄옵션 확인')
            url = 'http://cpsite.jclgift.com/'
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 분류, 인쇄옵션 확인 #!#!#!')


            driver.get(url)
            time.sleep(1)
            continue
        ###이미지 실패
        try:
            driver.find_element_by_name('opt_use_YN').click()
            driver.find_element_by_name('opt_use_YN').click()
            driver.find_element_by_name('opt_text_value').send_keys('100개 이하는 30,000원의 인쇄비가 추가됩니다.')

            #박스 옵션, 포장 옵션 상품 정보
            if '종이케이스' in all_values[i][29]:
                driver.find_elements_by_name('opt_use_9')[1].click()
                driver.find_element_by_name('opt_9_sub_2').send_keys('0')

            elif '자체케이스' in all_values[i][29]:
                driver.find_elements_by_name('opt_use_9')[1].click()
                driver.find_element_by_name('opt_9_sub_4').send_keys('0')
            else :
                driver.find_element_by_name('opt_use_9').click()

            if '종이케이스' in all_values[i][29]:
                driver.find_elements_by_name('opt_use_10')[1].click()
                driver.find_element_by_name('opt_10_sub_2').send_keys('300')

            else:
                driver.find_element_by_name('opt_use_10').click()

            driver.find_element_by_xpath('//select[@name="info_gubun"] /option[text()="기타"]').click()

            fullName =util.findImage(str(all_values[i][3]))
            driver.find_element_by_name('p_image_b').send_keys(fullName)

        except:
            exceptLine.append(int(i)+1)
            reason.append('이미지 이름 및 경로 확인')
            url = 'http://cpsite.jclgift.com/'
            driver.get(url)
            time.sleep(1)
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 이미지 이름 및 경로 확인 #!#!#!')
            continue
        ###상세 설명 실패


        try:
            driver.switch_to.frame(driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[15]/td[2]/iframe'))
            driver.find_element_by_xpath('//*[@id="smart_editor2_content"]/div[3]/ul/li[2]/button').click()
            driver.switch_to.frame(driver.find_element_by_xpath('//*[@id="se2_iframe"]'))
            driver.find_element_by_xpath('/html/body').send_keys(all_values[i][36])
            driver.find_element_by_xpath('/html/body').send_keys('\n\n')
            driver.find_element_by_xpath('/html/body').send_keys(str(all_values[i][31]).replace('>',' style="width: 770px; ">'))
            driver.find_element_by_xpath('/html/body').send_keys('\n\n')
            driver.find_element_by_xpath('/html/body').send_keys(str(all_values[i][32]).replace('>',' style="width: 770px; ">'))


            driver.switch_to_default_content()
            driver.switch_to.frame(driver.find_element_by_name('jobFrame'))

            driver.find_element_by_class_name('btn_write').click()
            time.sleep(2)
            driver.switch_to_window(driver.window_handles[0])
            driver.switch_to.frame(driver.find_element_by_name('jobFrame'))
            url = 'http://cpsite.jclgift.com/'
            driver.get(url)
        except:
            reason.append('이름 중복 또는 파일이 너무 큼 확인')
            exceptLine.append(int(i)+1)
            url = 'http://cpsite.jclgift.com/'
            print("#!#!#!#! "+str(all_values[i][0])+' / '+str(int(len(all_values))-2)+' 실패: 이름 중복 또는 파일이 너무 큼 확인 #!#!#!')
            driver.get(url)
            time.sleep(1)
            continue

        print("####### "+str(all_values[i][0])+' / '+ str(count) +' 완료 #####')

    util.saveExcel(exceptLine,reason)
