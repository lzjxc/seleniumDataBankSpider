#istore触点
istoreList = []
istoreButton = driver.find_element_by_xpath("//*[text()='istore小…']")
ActionChains(driver).move_to_element(istoreButton).move_by_offset(100, 0).click_and_hold().perform()
contentDetailList.append(driver.find_element_by_xpath(
    '//*[@id="main_container"]/div/div/div/div/div/div[1]/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div/div[2]/div/div[3]').text)
ActionChains(driver).move_to_element(istoreButton).move_by_offset(100, 0).click().perform()
sleep(2)

#istore细分
istoreDetail = ['手淘小程序', '支付宝小程序', '饿了么小程序']
for istore in istoreDetail:
    istoreXpath = "//*[text()=" + "'" + i + "'" + "]"
    ActionChains(driver).move_to_element_with_offset(driver.find_element_by_xpath(str(istoreXpath)), 20,
                                                     0).move_by_offset(0, -30).click_and_hold().perform()
    istoreButtonText = driver.find_element_by_css_selector(
        '#main_container > div > div > div > div > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)').text
    istoreList.append(istoreButtonText)
for i in istoreList:
    print(i)