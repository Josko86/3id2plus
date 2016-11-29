from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import time
from selenium import webdriver

driver = webdriver.Firefox()
time.sleep(5)
driver.quit()
# from selenium import webdriver
#
# firefox_capabilities = DesiredCapabilities.FIREFOX
# firefox_capabilities['marionette'] = True
# driver = webdriver.Firefox(capabilities=firefox_capabilities)