import codecs
import csv
import re
import time
from time import sleep

import xlsxwriter as xlsxwriter
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from xlwt import  *


import pymongo
from pymongo.errors import DuplicateKeyError


client = pymongo.MongoClient(host='192.168.0.47', port=27017)
db= client.DataBank

collection = db.Information
contentsTotal = []
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
chrome_driver = "C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe"

driver = webdriver.Chrome(chrome_driver, chrome_options=chrome_options)
print('===========================')
print(driver.title)
driver.implicitly_wait(10)
targetTime = driver.find_element_by_css_selector('#content_container > div:nth-child(1) > div.js-page-content.content___3G2Z0 > div > div.databank_component_ceilingBar.headerTimeBarCB___1lKUX > div > div > span.datepicker___2UcjI > div > span > input[type="text"]').get_attribute('value')
creatTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
secondMenu = driver.find_element_by_xpath("//*[text()='二级类目：']")
workbookTitle = 'databank_theWholeLink_'+targetTime+'.xlsx'
workbook = xlsxwriter.Workbook('databank_全链路分布_裙子新_'+targetTime+'.xlsx')
data_total = {}
data_total["_id"]="全链路分布"+creatTime
data_total["title"] = "adckid"
data_total["dateTime"] = creatTime





def insert_item(collection, item):
    try:
        collection.insert(dict(item))
    except DuplicateKeyError:
        print('日期重复')
        pass
    except Exception as e:
        print('error!')
        print(e)

def the_whole_link(item):
    #全链路分布
    #认知
    awareTotal = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[1]/div[2]/div[2]').text
    awareIncrease = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[1]/div[2]/ul/li[1]/span[2]').text
    awareDecrease = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[1]/div[2]/ul/li[2]/span[2]').text
    print('aware',awareTotal,awareIncrease,awareDecrease)


    #兴趣
    interestTotal = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[2]/div[2]/div[2]').text
    interestIncrease = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[2]/div[2]/ul/li[1]/span[2]').text
    interestDecrease = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[2]/div[2]/ul/li[2]/span[2]').text
    print('interest', interestTotal, interestIncrease, interestDecrease)

    #purchase
    purchaseTotal = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[3]/div[2]/div[2]').text
    activePurchaseConsumer = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[3]/div[2]/ul/li[1]').text
    purchaseIncrease = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[3]/div[2]/ul/li[2]/span[2]').text
    purchaseDecrease = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[3]/div[2]/ul/li[3]/span[2]').text
    print('purchase', purchaseTotal, activePurchaseConsumer, purchaseIncrease, purchaseDecrease)

    #loyalty
    loyaltyTotal = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[4]/div[2]/div[2]').text
    loyaltyIncrease = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[4]/div[2]/ul/li[1]/span[2]').text
    loyaltyDecrease = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[4]/div/div[2]/div/div[2]/div[2]/div/ul/li[4]/div[2]/ul/li[2]/span[2]').text
    print('loyalty', loyaltyTotal, loyaltyIncrease, loyaltyDecrease)


    the_whole_link_data = {
        'creatTime': creatTime,
        'time': targetTime,
        'awareTotal':awareTotal,
        'awareIncrease':awareIncrease,
        'awareDecrease':awareDecrease,
        'interestTotal':interestTotal,
        'interestIncrease':interestIncrease,
        'interestDecrease':interestDecrease,
        'purchaseTotal':purchaseTotal,
        'activePurchaseConsumer':activePurchaseConsumer,
        'purchaseIncrease':purchaseIncrease,
        'purchaseDecrease':purchaseDecrease,
        'loyaltyTotal':loyaltyTotal,
        'loyaltyIncrease':loyaltyIncrease,
        'loyaltyDecrease':loyaltyDecrease,
                           }
    data_total["the_whole_link_data"]=the_whole_link_data

    #insert_item(db.wholeLink,the_whole_link_data)

    the_whole_link_data_excel = (
        ['消费者全链路分布','人数'],
        ['awareTotal',awareTotal],
        ['awareIncrease',awareIncrease],
        ['awareDecrease',awareDecrease],
        ['interestTotal',interestTotal],
        ['interestIncrease',interestIncrease],
        ['interestDecrease',interestDecrease],
        ['purchaseTotal',purchaseTotal],
        ['activePurchaseConsumer',activePurchaseConsumer],
        ['purchaseIncrease',purchaseIncrease],
        ['purchaseDecrease',purchaseDecrease],
        ['loyaltyTotal',loyaltyTotal],
        ['loyaltyIncrease',loyaltyIncrease],
        ['loyaltyDecrease',loyaltyDecrease],
        ['creatTime', creatTime],
        ['time',targetTime],
    )

    worksheet = workbook.add_worksheet(item+'theWholeLinkData')
    row = 0
    col = 0

    for title, number in (the_whole_link_data_excel):
        worksheet.write(row,col, title)
        worksheet.write(row,col+1, number)
        row+=1


def touch_point_iframe(title,item):

    #认知互动触点分布
    driver.switch_to.frame(driver.find_element_by_xpath("//iframe[contains(@src,'insight-engine')]"))
    print(title + ' touch point starting')
    touch_point_iframe_dict = {}
    touch_point_iframe_dict.update(targetTime=targetTime)
    touch_point_iframe_dict.update(creatTime=creatTime)
    while not driver.find_elements_by_xpath("//*[text()='付费广告']"):
        driver.switch_to.default_content()
        print("未找到触点，刷新中")
        driver.refresh()
        '''
        ActionChains(driver).move_to_element(secondMenu).move_by_offset(100, 0).click().perform()
        sleep(1)
        ActionChains(driver).move_to_element(driver.find_element_by_xpath("//*[contains(text(), '" + item + "')]")).click().perform()
        '''
        sleep(1)
        driver.find_element_by_xpath("//*[text()='"+title+"']").click()
        sleep(1)
        driver.switch_to.frame(driver.find_element_by_xpath("//iframe[contains(@src,'insight-engine')]"))

    try:
        # 搜索
        point_css_selector = '#main_container > div > div > div > div > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div.react-grid-item.bi-widget.show-arrow.static.cssTransforms > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)'
        searchButton = driver.find_element_by_xpath("//*[text()='搜索']")
        ActionChains(driver).move_to_element(searchButton).move_by_offset(100,0).perform()
        sleep(1)
        search_button_text = driver.find_element_by_css_selector(point_css_selector).text
        content_num = re.findall(r"\d+\.?\d*\%", search_button_text)
        if len(content_num) > 0:
            touch_point_iframe_dict[title + '搜索' + '_' + '细分'] = content_num[0]
            touch_point_iframe_dict[title + '搜索' + '_' + 'top5'] = content_num[1]
        ActionChains(driver).move_to_element(searchButton).move_by_offset(100, 0).click().perform()
        print('search ok')

        #payedAd触点
        payedAdList = []

        ActionChains(driver).move_by_offset(400,400).perform()
        sleep(1)
        ActionChains(driver).move_to_element(driver.find_element_by_xpath("//*[text()='线下触点']")).move_by_offset(100, 0).click().perform()
        print("adButton clicked")
        sleep(1)
        ActionChains(driver).move_to_element(driver.find_element_by_xpath("//*[text()='付费广告']")).move_by_offset(100, 0).perform()
        print("adButtonText ok")
        sleep(1)
        payedAd_button_text = driver.find_element_by_css_selector(point_css_selector).text
        print("get adButtonText ok")
        print(payedAd_button_text)

        sleep(1)
        print('movement ok')
        print("adButton clicked")
        ActionChains(driver).move_to_element(driver.find_element_by_xpath("//*[text()='付费广告']")).move_by_offset(100,
                                                                                                                0).click().perform()
        sleep(1)
        content_num = re.findall(r"\d+\.?\d*\%", payedAd_button_text)
        print("regular ok")
        if len(content_num) > 0:
            touch_point_iframe_dict[title + '付费广告' + '_' + '细分'] = content_num[0]
            touch_point_iframe_dict[title + '付费广告' + '_' + 'top5'] = content_num[1]
        sleep(1)
        print('payedad ok')

        #payedAd细分
        payedAdDetail = ['Uni Desk', '优酷广告', '一夜霸屏', '品牌雷达', '品牌专区', '明星店铺', '钻石展位', '品牌特秀', '摇一摇', '事件营销']
        for payedAd in payedAdDetail:
            touch_point_iframe_dict[title+'付费广告'+'_'+payedAd+'_'+'细分'] =''
            touch_point_iframe_dict[title + '付费广告' + '_' + payedAd + '_' + 'top5']=''
            while touch_point_iframe_dict[title+'付费广告'+'_'+payedAd+'_'+'细分'] == '':
                payedAdXpath = "//*[text()='" + payedAd + "']"
                ActionChains(driver).move_to_element_with_offset(driver.find_element_by_xpath(payedAdXpath), 30,
                                                                 0).move_by_offset(0, -30).perform()
                payedAdButtonText = driver.find_element_by_css_selector(
                    '#main_container > div > div > div > div.react-grid-item.bi-widget.active.static.cssTransforms > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)').text
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '#main_container > div > div > div > div.react-grid-item.bi-widget.active.static.cssTransforms > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)')))
                content_num = re.findall(r"\d+\.?\d*\%", payedAdButtonText)
                if len(content_num) > 0:
                    touch_point_iframe_dict[title+'付费广告'+'_'+payedAd+'_'+'细分']=content_num[0]
                    touch_point_iframe_dict[title+'付费广告'+'_'+payedAd+'_'+'top5']=content_num[1]




        #内容运营
        contentMarketingButton = driver.find_element_by_xpath("//*[text()='内容运营']")
        ActionChains(driver).move_to_element(contentMarketingButton).move_by_offset(100,0).perform()
        contentMarketingButton_text = driver.find_element_by_xpath('//*[@id="main_container"]/div/div/div/div/div/div[1]/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div/div[2]/div/div[3]').text
        ActionChains(driver).move_to_element(contentMarketingButton).move_by_offset(100,0).click().perform()
        content_num = re.findall(r"\d+\.?\d*\%", contentMarketingButton_text)
        if len(content_num) > 0:
            touch_point_iframe_dict[title + '内容运营' + '_' + '细分'] = content_num[0]
            touch_point_iframe_dict[title + '内容运营' + '_' + 'top5'] = content_num[1]

        #内容运营细分
        contentDetail = ['品牌号','淘宝头条','有好货','必买清单','猜你喜欢','生活研究所','直播','微淘','淘宝短视频','每日好店','运动俱乐部']
        for i in contentDetail:
            touch_point_iframe_dict[title + '内容运营' + '_' + i + '_' + '细分'] = ''
            touch_point_iframe_dict[title + '内容运营' + '_' + i + '_' + 'top5'] = ''
            while touch_point_iframe_dict[title + '内容运营' + '_' + i + '_' + '细分'] == '':
                ixpath =  "//*[text()="+"'"+i+"'"+"]"
                ActionChains(driver).move_to_element_with_offset(driver.find_element_by_xpath(str(ixpath)),30,0).move_by_offset(0,-30).perform()
                iButtonText = driver.find_element_by_css_selector('#main_container > div > div > div > div.react-grid-item.bi-widget.active.static.cssTransforms > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)').text
                content_num = re.findall(r"\d+\.?\d*\%", iButtonText)
                if len(content_num) > 0:
                    touch_point_iframe_dict[title + '内容运营' + '_' + i + '_' + '细分'] = content_num[0]
                    touch_point_iframe_dict[title + '内容运营' + '_' + i + '_' + 'top5'] = content_num[1]


        #tianMaoMarketing触点
        tianMaoMarketingButton = driver.find_element_by_xpath("//*[text()='天猫营销…']")
        ActionChains(driver).move_to_element(tianMaoMarketingButton).move_by_offset(100, 0).perform()
        tianMaoMarketingButton_text = driver.find_element_by_xpath('//*[@id="main_container"]/div/div/div/div/div/div[1]/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div/div[2]/div/div[3]').text
        ActionChains(driver).move_to_element(tianMaoMarketingButton).move_by_offset(100, 0).click().perform()
        content_num = re.findall(r"\d+\.?\d*\%", tianMaoMarketingButton_text)
        if len(content_num) > 0:
            touch_point_iframe_dict[title + '天猫营销平台' + '_' + '细分'] = content_num[0]
            touch_point_iframe_dict[title + '天猫营销平台' + '_' + 'top5'] = content_num[1]
        #tianMaoMarketing细分
        contentDetail = ['超级品牌日', '互动吧', '天合&流量宝', '聚划算', '试用中心', '天猫U先', '淘抢购', '欢聚日', '全明星计划', '天猫新人礼']
        for i in contentDetail:
            touch_point_iframe_dict[title + '天猫营销平台' + '_' + i + '_' + '细分'] = ''
            touch_point_iframe_dict[title + '天猫营销平台' + '_' + i + '_' + 'top5'] = ''
            while touch_point_iframe_dict[title + '天猫营销平台' + '_' + i + '_' + '细分'] == '':
                ixpath = "//*[text()=" + "'" + i + "'" + "]"
                ActionChains(driver).move_to_element_with_offset(driver.find_element_by_xpath(str(ixpath)),30,0).move_by_offset(0,
                                                                                                               -30).perform()
                iButtonText = driver.find_element_by_css_selector(
                    '#main_container > div > div > div > div.react-grid-item.bi-widget.active.static.cssTransforms > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)').text
                content_num = re.findall(r"\d+\.?\d*\%", iButtonText)
                if len(content_num) > 0:
                    touch_point_iframe_dict[title + '天猫营销平台' + '_' + i + '_' + '细分'] = content_num[0]
                    touch_point_iframe_dict[title + '天猫营销平台' + '_' + i + '_' + 'top5'] = content_num[1]


        #marketingChannel触点

        marketingChannelButton = driver.find_element_by_xpath("//*[text()='销售渠道']")
        ActionChains(driver).move_to_element(marketingChannelButton).move_by_offset(100, 0).perform()
        marketingChannelButton_text = driver.find_element_by_xpath(
            '//*[@id="main_container"]/div/div/div/div/div/div[1]/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div/div[2]/div/div[3]').text
        ActionChains(driver).move_to_element(marketingChannelButton).move_by_offset(100, 0).click().perform()
        content_num = re.findall(r"\d+\.?\d*\%", marketingChannelButton_text)
        if len(content_num) > 0:
            touch_point_iframe_dict[title + '销售渠道' + '_' + '细分'] = content_num[0]
            touch_point_iframe_dict[title + '销售渠道' + '_' + 'top5'] = content_num[1]

        # marketingChannel细分
        marketing_channel_tianmao_element = driver.find_element_by_xpath("//*[text()='智慧门店']/../..")
        tianmao_elements = marketing_channel_tianmao_element.find_elements_by_tag_name('text')
        for tianmao_element in tianmao_elements:
            if '%' not in tianmao_element.text:
                ActionChains(driver).move_to_element_with_offset(tianmao_element, 50, 0).move_by_offset(0,
                                                                                                -30).perform()
                tianmao_element_content = driver.find_element_by_css_selector(
                    "#main_container > div > div > div > div.react-grid-item.bi-widget.active.static.cssTransforms > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)").text
                content_num = re.findall(r"\d+\.?\d*\%", tianmao_element_content)
                if len(content_num) > 0:
                    touch_point_iframe_dict[title + '天猫营销平台' + '_' + tianmao_element.text + '_' + '细分'] = content_num[0]
                    touch_point_iframe_dict[title + '天猫营销平台' + '_' + tianmao_element.text + '_' + 'top5'] = content_num[1]

        #offlineTouchSpot触点

        offlineTouchSpotButton = driver.find_element_by_xpath("//*[text()='线下触点']")
        ActionChains(driver).move_to_element(offlineTouchSpotButton).move_by_offset(100, 0).perform()
        offlineTouchSpotButton_text = driver.find_element_by_xpath(
            '//*[@id="main_container"]/div/div/div/div/div/div[1]/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div/div[2]/div/div[3]').text
        ActionChains(driver).move_to_element(offlineTouchSpotButton).move_by_offset(100, 0).click().perform()
        content_num = re.findall(r"\d+\.?\d*\%", offlineTouchSpotButton_text)
        if len(content_num) > 0:
            touch_point_iframe_dict[title + '线下触点' + '_' + '细分'] = content_num[0]
            touch_point_iframe_dict[title + '线下触点' + '_' + 'top5'] = content_num[1]

        #offlineTouchSpot细分
        offlineTouchSpotDetail = ['菜鸟驿站', '天猫U先', '智慧门店', '智慧商圈', '淘鲜达', '淘宝彩蛋', '智能母婴室']
        for offlineTouch in offlineTouchSpotDetail:
            touch_point_iframe_dict[title + '天猫营销平台' + '_' + offlineTouch + '_' + '细分'] = ''
            touch_point_iframe_dict[title + '天猫营销平台' + '_' + offlineTouch + '_' + 'top5'] = ''
            while touch_point_iframe_dict[title + '天猫营销平台' + '_' + offlineTouch + '_' + '细分'] == '':
                offlineTouchXpath = "//*[text()=" + "'" + offlineTouch + "'" + "]"
                ActionChains(driver).move_to_element(driver.find_element_by_xpath(str(offlineTouchXpath))).move_by_offset(0,
                                                                                                                          -30).perform()
                offlineTouchButtonText = driver.find_element_by_css_selector(
                    '#main_container > div > div > div > div.react-grid-item.bi-widget.active.static.cssTransforms > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)').text
                content_num = re.findall(r"\d+\.?\d*\%", offlineTouchButtonText)
                if len(content_num) > 0:
                    touch_point_iframe_dict[title + '天猫营销平台' + '_' + offlineTouch + '_' + '细分'] = content_num[0]
                    touch_point_iframe_dict[title + '天猫营销平台' + '_' + offlineTouch + '_' + 'top5'] = content_num[1]


        #istore触点

        istoreButton = driver.find_element_by_xpath("//*[text()='istore小…']")
        ActionChains(driver).move_to_element(istoreButton).move_by_offset(100, 0).perform()
        istoreButton_text = driver.find_element_by_xpath(
            '//*[@id="main_container"]/div/div/div/div/div/div[1]/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div/div[2]/div/div[3]').text
        ActionChains(driver).move_to_element(istoreButton).move_by_offset(100, 0).click().perform()
        content_num = re.findall(r"\d+\.?\d*\%", istoreButton_text)
        if len(content_num) > 0:
            touch_point_iframe_dict[title + 'istore小程序' + '_' + '细分'] = content_num[0]
            touch_point_iframe_dict[title + 'istore小程序' + '_' + 'top5'] = content_num[1]

        #istore细分
        istoreDetail = ['手淘小程序', '支付宝小程序', '饿了么小程序']
        for istore in istoreDetail:
            touch_point_iframe_dict[title + '天猫营销平台' + '_' + istore + '_' + '细分'] = ''
            touch_point_iframe_dict[title + '天猫营销平台' + '_' + istore + '_' + 'top5'] = ''
            while touch_point_iframe_dict[title + '天猫营销平台' + '_' + istore + '_' + '细分'] == '':
                istoreXpath = "//*[text()=" + "'" + istore + "'" + "]"
                ActionChains(driver).move_to_element_with_offset(driver.find_element_by_xpath(str(istoreXpath)), 20,
                                                                 0).move_by_offset(0, -30).perform()
                istoreButtonText = driver.find_element_by_css_selector(
                    '#main_container > div > div > div > div.react-grid-item.bi-widget.active.static.cssTransforms > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)').text
                content_num = re.findall(r"\d+\.?\d*\%", istoreButtonText)
                if len(content_num) > 0:
                    touch_point_iframe_dict[title + '天猫营销平台' + '_' + istore + '_' + '细分'] = content_num[0]
                    touch_point_iframe_dict[title + '天猫营销平台' + '_' + istore + '_' + 'top5'] = content_num[1]


        #插入数据库
        #insert_item(db.touchPoint,touch_point_iframe_dict)
        data_total[title+'touch point'] = touch_point_iframe_dict

        #写入excel
        worksheet = workbook.add_worksheet(item + title+'touch point')
        row = 0
        col = 0
        touch_point_iframe_list = []
        for k, v in touch_point_iframe_dict.items():
            print(k, v)
            tempList = [k, v]
            touch_point_iframe_list.append(tempList)
        print(touch_point_iframe_list)
        touch_point_iframe_tuple = tuple(touch_point_iframe_list)
        print(touch_point_iframe_tuple)

        for title, number in (touch_point_iframe_tuple):
            worksheet.write(row, col, title)
            worksheet.write(row, col + 1, number)
            row += 1

    except NoSuchElementException as nse:
        print('找不到触点元素')
    except Exception as e:
        pass

    print("touch point finished")
    driver.switch_to.default_content()

def search_iframe(item):
    driver.switch_to.frame(driver.find_element_by_xpath("//iframe[contains(@src,'insight-engine')]"))
    while not driver.find_elements_by_xpath("//span[@data-word='abckids童装']/../span[position()<11]"):
        driver.switch_to.default_content()
        print("未找到搜索词，刷新中")
        driver.refresh()
        '''
        ActionChains(driver).move_to_element(secondMenu).move_by_offset(100, 0).click().perform()
        sleep(1)
        ActionChains(driver).move_to_element(
            driver.find_element_by_xpath("//*[contains(text(), '" + item + "')]")).click().perform()
        '''
        sleep(1)

        driver.find_element_by_xpath("//*[text()='兴趣']").click()
        sleep(10)
        driver.switch_to.frame(driver.find_element_by_xpath("//iframe[contains(@src,'insight-engine')]"))
    search_iframe_dict = {}

    try:
        search_keyword_list = []
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[@data-word='abckids童装']")))
        search_words = driver.find_elements_by_xpath("//span[@data-word='abckids童装']/../span[position()<11]")
        for i in search_words:

            #top50搜索词
            ActionChains(driver).move_to_element(i).click().perform()
            sleep(1)
            ActionChains(driver).move_by_offset(300,300).move_to_element(driver.find_element_by_xpath("//span[@data-word='"+i.text+"']")).perform()
            iContent = driver.find_element_by_xpath("//*[contains(text(), '品牌搜索词搜索次数占比:')]").text
            iName = driver.find_element_by_xpath("//*[contains(text(), '品牌搜索词搜索次数占比:')]/preceding-sibling::div[1]").text
            print(iName,iContent)
            content_num = re.findall(r"\d+\.?\d*\%", iContent)
            search_iframe_dict[iName]=content_num[0]



            #上游词
            print("-" * 20)
            sleep(1)
            try:
                up_words = driver.find_element_by_xpath("//*[text()='没有上游词']/../..")
                up_word = up_words.find_elements_by_tag_name('text')
                for uword in up_word:
                    ActionChains(driver).move_to_element(uword).move_by_offset(100,0).perform()
                    uword_content = driver.find_element_by_css_selector("#main_container > div > div > div > div:nth-child(2) > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)").text
                    print(uword_content)
                    content_name = re.match(r'.+',uword_content).group()
                    content_num = re.findall(r"\d+\.?\d*\%", uword_content)
                    content_name_clean = content_name.replace(".","point")
                    search_iframe_dict[iName + '_' + content_name_clean] = content_num[0]
            except:
                print("no up word")
                continue

            #下游词
            try:
                print("-" * 20)
                down_words = driver.find_element_by_xpath("//*[text()='没有下游词']/../..")
                down_word = down_words.find_elements_by_tag_name('text')
                for dword in down_word:
                    ActionChains(driver).move_to_element(dword).move_by_offset(100,0).perform()
                    dword_content = driver.find_element_by_css_selector("#main_container > div > div > div > div:nth-child(2) > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(1) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)").text
                    print(dword_content)
                    content_name = re.match(r'.+', dword_content).group()
                    content_name_clean = content_name.replace(".","point")
                    content_num = re.findall(r"\d+\.?\d*\%", dword_content)
                    search_iframe_dict[iName + '_' + content_name_clean] = content_num[0]
            except:
                print("no down word")
                continue

        # 插入数据库
        #insert_item(db.searchIframe, search_iframe_dict)
        data_total[item+'品牌搜索词']=search_iframe_dict

        # 写入excel
        worksheet = workbook.add_worksheet(item + '品牌搜索词')
        row = 0
        col = 0
        search_iframe_list = []
        for k, v in search_iframe_dict.items():
            print(k, v)
            tempList = [k, v]
            search_iframe_list.append(tempList)
        print(search_iframe_list)
        search_iframe_tuple = tuple(search_iframe_list)
        print(search_iframe_tuple)

        for title, number in (search_iframe_tuple):
            worksheet.write(row, col, title)
            worksheet.write(row, col + 1, number)
            row += 1

    except NoSuchElementException as nse :
        print("can not find the searching word")

    except Exception as e:
        print('search_iframe error raised')
    print("search frame finished")
    driver.switch_to.default_content()

def purchase_foot_print(item):

    driver.switch_to.frame(driver.find_element_by_xpath("//iframe[contains(@src,'insight-engine')]"))

    while not driver.find_elements_by_xpath("//*[text()='<=1天']"):
        driver.switch_to.default_content()
        print("未找到购买足迹，刷新中")
        driver.refresh()
        '''
        ActionChains(driver).move_to_element(secondMenu).move_by_offset(100, 0).click().perform()
        sleep(1)
        ActionChains(driver).move_to_element(
            driver.find_element_by_xpath("//*[contains(text(), '" + item + "')]")).click().perform()
        '''
        sleep(1)
        driver.find_element_by_xpath("//*[text()='购买']").click()
        sleep(10)
        driver.switch_to.frame(driver.find_element_by_xpath("//iframe[contains(@src,'insight-engine')]"))

    try:
        purchase_foot_print_dict = {}
        time_lag_list = ['<=1天','2~7天','8~30天','1~2个月','2~4个月','4~6个月','6~8个月','8~10个月','10~12个月','>1年']
        purchasing_channels = ['天猫国际旗舰店', '天猫国际其他', '天猫国际直营', '天猫超市', '天猫旗舰店', '天猫其他', '全球购', '淘宝集市']
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[text()='<=1天']")))
        for time_lag in range(len(time_lag_list)):
            ActionChains(driver).move_to_element(driver.find_element_by_xpath("//*[text()='" + time_lag_list[time_lag] + "']")).move_by_offset(100,0).perform()
            tagcontent = driver.find_element_by_css_selector("#main_container > div > div > div > div:nth-child(1) > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div.react-grid-item.bi-widget.show-arrow.static.cssTransforms > div > div.default-wrapper-content > div.react-grid-layout.group-container-content > div:nth-child(2) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)").text
            content_name = re.match(r'.+', tagcontent).group()
            content_num = re.findall(r"\d+\.?\d*\%", tagcontent)
            purchase_foot_print_dict[content_name + '_购买消费者'] = content_num[0]
            purchase_foot_print_dict[content_name + '_top5'] = content_num[1]
            ActionChains(driver).move_to_element(driver.find_element_by_xpath(str("//*[text()='" + time_lag_list[time_lag] + "']"))).move_by_offset(100,0).click().perform()
            if time_lag != 0:
                ActionChains(driver).move_to_element(driver.find_element_by_xpath(str("//*[text()='" + time_lag_list[time_lag -1] + "']"))).move_by_offset(100,0).click().perform()
            for channel in purchasing_channels:
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[text()='" + channel + "']")))
                    ActionChains(driver).move_to_element_with_offset(driver.find_element_by_xpath("//*[text()='" + channel + "']"), 30,0).move_by_offset(0,-30).perform()
                    channel_content = driver.find_element_by_css_selector("#main_container > div > div > div > div:nth-child(1) > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)").text
                    print(channel_content)
                    content_name = re.match(r'.+', channel_content).group()
                    content_num = re.findall(r"\d+\.?\d*\%", channel_content)
                    purchase_foot_print_dict[content_name + '-' + channel + '_购买消费者'] = content_num[0]
                    purchase_foot_print_dict[content_name + '-' + channel + '_top5'] = content_num[1]
                except:
                    print('purchase_foot_print wait too long')
                    pass

        # 插入数据库
        #insert_item(db.purchaseFootPrint, purchase_foot_print_dict)
        data_total[item+"_foot_print"] = purchase_foot_print_dict

        # 写入excel
        worksheet = workbook.add_worksheet(item + '购买足迹分析')
        row = 0
        col = 0
        purchase_foot_print_list = []
        for k, v in purchase_foot_print_dict.items():
            print(k, v)
            tempList = [k, v]
            purchase_foot_print_list.append(tempList)
        print(purchase_foot_print_list)
        purchase_foot_print_tuple = tuple(purchase_foot_print_list)
        print(purchase_foot_print_tuple)

        for title, number in (purchase_foot_print_tuple):
            worksheet.write(row, col, title)
            worksheet.write(row, col + 1, number)
            row += 1

    except NoSuchElementException as nse:
        print('no purchase_foot_print element')
    except Exception as e:
        print('purchase_foot_print error appear')

    print("purchase foot point finished")
    driver.switch_to.default_content()

def loyal_foot_point(item):
    driver.switch_to.frame(driver.find_element_by_xpath("//iframe[contains(@src,'insight-engine')]"))
    print("start loyal foot point")

    while not driver.find_elements_by_xpath("//*[text()='2天']"):
        driver.switch_to.default_content()
        print("未找到复购足迹，刷新中")
        driver.refresh()
        '''
        ActionChains(driver).move_to_element(secondMenu).move_by_offset(100, 0).click().perform()
        sleep(1)
        ActionChains(driver).move_to_element(
            driver.find_element_by_xpath("//*[contains(text(), '" + item + "')]")).click().perform()
        '''
        sleep(2)
        driver.find_element_by_xpath("//*[text()='忠诚']").click()
        sleep(10)
        driver.switch_to.frame(driver.find_element_by_xpath("//iframe[contains(@src,'insight-engine')]"))
    print("while ok")
    try:
        re_purchase_time=['2天','3天', '4天', '5天', '6天', '7天', '8天', '9天', '10天', '超过10天']
        re_purchase_period = ['1天','2-7天', '8-30天','1-2个月','2-4个月','4-6个月','6-8个月','8-10个月','10-12个月']
        purchase_channels = ['天猫国际旗舰店', '天猫国际其他', '天猫国际直营', '天猫超市', '天猫旗舰店', '天猫其他', '全球购', '淘宝集市']

        print("try ok")
        #最近一年复购天数
        recent_one_year_data=[]
        loyal_foot_point_list=[]
        for word in re_purchase_time:
            loyal_time_content =''
            while loyal_time_content == '':
                ActionChains(driver).move_to_element_with_offset(
                    driver.find_element_by_xpath(str("//*[text()='" + word + "']")), 20, 0).move_by_offset(0, -100).perform()
                print('date found')
                loyal_time_content = driver.find_element_by_css_selector('#main_container > div > div > div > div:nth-child(1) > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(2) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)').text
                print("date content found",loyal_time_content)
                if len(loyal_time_content) > 1:
                    word_content_name = re.match(r'.+', loyal_time_content).group()
                    word_content_num = re.findall(r"\d+\.?\d*\%", loyal_time_content)
                    print("regular ok")
                    dict_tem = {'name_re_purchase_days': word_content_name + '_复购消费者',
                                'value_re_purchase_days': word_content_num[0],
                                'name_competitor_days': word_content_name + '_同行业竞争品牌平均',
                                'value_competitor_days': word_content_num[1]}
                    for k, v in dict_tem.items():
                        print(k, v)
                        tempList = [k, v]
                        loyal_foot_point_list.append(tempList)
                    recent_one_year_data.append(dict_tem)

        #复购周期分布
        re_purchase_period_data=[]
        for period in re_purchase_period:
            ActionChains(driver).move_to_element_with_offset(driver.find_element_by_xpath(str("//*[text()='" + period + "']")), 20 , 0).move_by_offset(0,-150).perform()
            loyal_period_content = driver.find_element_by_css_selector('#main_container > div > div > div > div:nth-child(1) > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(1) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)').text
            period_content_name = re.match(r'.+', loyal_period_content).group()
            period_content_num = re.findall(r"\d+\.?\d*\%", loyal_period_content)
            dict_tem = {'name_re_purchase_period': period_content_name + '_复购消费者',
                        'value_re_purchase_period': period_content_num[0],
                        'name_competitor_period': period_content_name + '_同行业竞争品牌平均',
                        'value_competitor_period': period_content_num[1]}
            for k, v in dict_tem.items():
                print(k, v)
                tempList = [k, v]
                loyal_foot_point_list.append(tempList)
            re_purchase_period_data.append(dict_tem)

        #购买渠道分布
        purchase_channels_data=[]
        for channel in purchase_channels:
            ActionChains(driver).move_to_element(
                driver.find_element_by_xpath(str("//*[text()='" + channel + "']"))).move_by_offset(0,
                                                                                                                      -150).perform()
            channel_content = driver.find_element_by_css_selector(
                '#main_container > div > div > div > div:nth-child(1) > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(6) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)').text
            channel_content_name = re.match(r'.+', channel_content).group()
            channel_content_num = re.findall(r"\d+\.?\d*\%", channel_content)
            dict_tem = {'name_this_channel_percent': channel_content_name + '_复购消费者在该渠道购买过占比',
                        'value_this_channel_percent': channel_content_num[0],
                        'name_top5_competitor_percent': channel_content_name + '_同行业TOP5品牌平均占比',
                        'value_top5_competitor_percent': channel_content_num[1],
                        'name_only_this_channel_percent': channel_content_name + '_复购消费者仅在该渠道购买过占比',
                        'value_only_this_channel_percent': channel_content_num[2],
                        'name_top5_competitor_percent_sec': channel_content_name + '_同行业TOP5品牌平均占比',
                        'value_top5_competitor_percent_sec': channel_content_num[3]
                        }
            for k, v in dict_tem.items():
                print(k, v)
                tempList = [k, v]
                loyal_foot_point_list.append(tempList)
            purchase_channels_data.append(dict_tem)

        # 插入数据库
        #insert_item(db.loyalFootPoint, loyal_foot_point_dict)
        loyal_foot_point_dict={'recent_one_year_data':recent_one_year_data,
                               're_purchase_period_data':re_purchase_period_data,
                               'purchase_channels_data':purchase_channels_data}
        data_total[item+'_loyal_foot_point'] = loyal_foot_point_dict

        # 写入excel
        worksheet = workbook.add_worksheet(item+'复购足迹分析')
        row = 0
        col = 0
        #for k, v in loyal_foot_point_dict.items():
         #   print(k, v)
          #  tempList = [k, v]
           # loyal_foot_point_list.append(tempList)
        print(loyal_foot_point_dict)
        loyal_foot_point_tuple = tuple(loyal_foot_point_list)

        for title, number in (loyal_foot_point_tuple):
            worksheet.write(row, col, title)
            worksheet.write(row, col + 1, number)
            row += 1

    except NoSuchElementException as nse:
        print('no  loyal_foot_point element')
    except Exception:
        print('loyal_foot_point error appear')

    driver.switch_to.default_content()

def crawlAll(item):

    # 购买tag
    #driver.find_element_by_xpath("//*[text()='购买']").click()
    #sleep(3)
    purchase_foot_print(item)
    touch_point_iframe('购买', item)

    # 忠诚tag
    driver.find_element_by_xpath("//*[text()='忠诚']").click()
    sleep(3)
    loyal_foot_point(item)
    touch_point_iframe('忠诚',item)

    # 兴趣tag
    driver.find_element_by_xpath("//*[text()='兴趣']").click()
    sleep(3)
    touch_point_iframe('兴趣', item)
    search_iframe(item)

    # 认知tag
    driver.find_element_by_xpath("//*[text()='认知']").click()
    sleep(3)
    the_whole_link(item)
    touch_point_iframe('认知', item)


def main():
    try:
        '''
        ActionChains(driver).move_to_element(secondMenu).move_by_offset(100,0).click().perform()
        menu = driver.find_elements_by_xpath('/html/body/div[3]/div/div/ul/li')
        menuList = []
        for i in menu:
            if i.text not in menuList:
                menuList.append(i.text)
        ActionChains(driver).move_to_element(secondMenu).move_by_offset(100, 0).click().perform()
        for item in menuList:
            #sleep(3)
            ActionChains(driver).move_to_element(secondMenu).move_by_offset(100, 0).click().perform()
            #sleep(2)
            ActionChains(driver).move_to_element(driver.find_element_by_xpath("//*[text()='" + item + "']")).click().perform()
            #sleep(2)
            print(item)
            if '>' in item:
                item_name = str(re.findall(r"\>.+", item))
                name = str(re.findall(r"[\u4e00-\u9fa5]+",item_name)[0])
                crawlAll(name)
            else:
                crawlAll(item)
            print("changes page")
            '''
        crawlAll('All_category')
        insert_item(db.dataTotal, data_total)

    finally:
        workbook.close()
        pass


if __name__=='__main__':
    main()



