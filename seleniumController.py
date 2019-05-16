import os

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import ActionChains
import time

chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
chrome_driver = "C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe"

driver = webdriver.Chrome(chrome_driver, chrome_options=chrome_options)
print('===========================')
print(driver.title)

activeConsumer = driver.find_element_by_css_selector('#content_container > div:nth-child(1) > div.sidebar___BsOtq > ul > li.next-navigation-item.next-navigation-item-selected.next-navigation-item-selected-left.item___3eJxg > div > div > div > span')
print(activeConsumer.text)


#自定义分析
allLink = driver.find_element_by_xpath('//*[@id="content_container"]/div[1]/div[3]/ul/li[5]/div/div/div/span[1]')
allLink.click()

time.sleep(3)
#新建自定义人群
createCustomPeople = driver.find_element_by_css_selector('#content_container > div:nth-child(1) > div.js-page-content.content___3G2Z0 > div > div.next-tabs-content > div.next-tabs-tabpane.active > div > div:nth-child(1) > div.operation___1ou_- > div:nth-child(2) > button')
createCustomPeople.click()

#people input test
time.sleep(5)
print('===========================')
print(driver.title)

#获取当前窗口句柄
windowHandle = driver.current_window_handle
windowHandleAll = driver.window_handles
print("----------")
print('windowHandle')
for handle in windowHandleAll:
    print(handle)

#人群名称
driver.switch_to.window(windowHandleAll[1])
peopleInput = driver.find_element_by_css_selector('#content_container > div:nth-child(1) > div.js-page-content.content___3G2Z0 > div > div > div:nth-child(2) > div > div.src-component-topInfoBar-crowdBox-UORZk > span > input[type="text"]')
peopleInput.clear()
peopleInput.send_keys('123')

#添加店铺商品圈人

'''try:
    storePeople = driver.find_element_by_css_selector('#crowdPickMainArea > div.src-sideBar-2qHts > div > div:nth-child(3) > div.src-component-PickPatternMenu-components-PickPatternGroup-itemsArea-UGskf > div > div.src-component-PickPatternMenu-components-PickPatternItem-label-1a-M7')
    controlPad = driver.find_element_by_css_selector('#crowdPickConditionArea > div')
    allLink = driver.find_element_by_css_selector('#crowdPickMainArea > div.src-sideBar-2qHts > div > div:nth-child(1) > div.src-component-PickPatternMenu-components-PickPatternGroup-itemsArea-UGskf > span > div:nth-child(1)')
    actions = ActionChains(driver)
    actions.drag_and_drop_by_offset(storePeople,100,0)
    actions.perform()

except Exception as e:
    print(e)
finally:
    print('finished')'''

with open(os.path.abspath('drag_and_drop_helper.js'), 'r') as js_file:
    line = js_file.readline()
    script = ''
    while line:
        script += line
        line = js_file.readline()

time.sleep(3)
print(driver.title)
driver.execute_script(script + "$('#src-component-PickPatternMenu-components-PickPatternItem-wrapper-1dr1x').simulateDragDrop({ dropTarget: '#crowdPickConditionArea'});")