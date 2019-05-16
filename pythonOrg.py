from selenium import webdriver
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome()
driver.get("https://login.taobao.com/?style=mini&full_redirect=true&newMini2=true&from=databank&sub=true&redirectURL=https://databank.tmall.com/")
elemButton = driver.find_element_by_xpath('//*[@id="J_Quick2Static"]')
elemButton.click()
elem = driver.find_element_by_id("TPL_username_1")
elem.clear()
elem.send_keys("pycon")
elem.send_keys(Keys.RETURN)
assert "No results found." not in driver.page_source
