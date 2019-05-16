#payedAd触点
payedAdList = []
payedAdButton = driver.find_element_by_xpath("//*[text()='付费广告']")
ActionChains(driver).move_to_element(payedAdButton).move_by_offset(100, 0).click_and_hold().perform()
contentDetailList.append(driver.find_element_by_xpath(
    '//*[@id="main_container"]/div/div/div/div/div/div[1]/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div/div[2]/div/div[3]').text)
ActionChains(driver).move_to_element(payedAdButton).move_by_offset(100, 0).click().perform()
sleep(2)

#payedAd细分
payedAdDetail = ['Uni Desk', '优酷广告', '一夜霸屏', '品牌雷达', '品牌专区', '明星店铺', '钻石展位', '品牌特秀', '摇一摇', '事件营销']
for payedAd in payedAdDetail:
    payedAdXpath = "//*[text()=" + "'" + payedAd + "'" + "]"
    ActionChains(driver).move_to_element_with_offset(driver.find_element_by_xpath(str(payedAdXpath)), 20,
                                                     0).move_by_offset(0, -30).click_and_hold().perform()
    payedAdButtonText = driver.find_element_by_css_selector(
        '#main_container > div > div > div > div > div > div.ysf-wrapper > div:nth-child(2) > div > div.react-grid-layout.group-container-content > div:nth-child(3) > div > div.default-wrapper-content > div.ysf-chart-container.full-screen.enhance-bar-click-wrap > div > div.content > div > div:nth-child(4)').text
    payedAdList.append(payedAdButtonText)
for i in payedAdList:
    print(i)
contentsTotal.append(payedAdList)