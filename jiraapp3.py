from selenium import webdriver

driver = webdriver.Chrome('C:/Users/kprokopiuk/Downloads/chromedriver.exe')
driver.get('https://partnerjira.g2-networks.com/secure/Dashboard.jspa?selectPageId=16395')
element = driver.find_element_by_css_selector('a.aui-nav-link.login-link')
element.click()
element = driver.find_element_by_id('login-form-username')
element.send_keys('login')
element = driver.find_element_by_id('login-form-password')
element.send_keys('password')
element = driver.find_element_by_id('login-form-submit')
element.click()
element = driver.find_element_by_id('create_link')
element.click()

# create jira ticket in browser
