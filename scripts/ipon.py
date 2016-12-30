import os

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import time
import win32com.client



def set_up_browser():
    # set up browser
    # Windows
    if os.name == 'nt':
        ############################# FIREFOX ###########################################
        binary = FirefoxBinary(r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe')
        profile = webdriver.FirefoxProfile(r'C:\Users\josko\AppData\Roaming\Mozilla\Firefox\Profiles\b6lwaicj.default')
        capabilities = webdriver.DesiredCapabilities().FIREFOX
        capabilities['acceptSslCerts'] = True
        profile.accept_untrusted_certs = True
        browser = webdriver.Firefox(firefox_binary=binary, firefox_profile=profile, capabilities=capabilities)
        browser.capabilities['acceptSslCerts'] = True
    browser.get('https://lyon.metagate.orange.com/dana/home/index.cgi')
    return browser

def login(browser):
    # login
    elem = browser.find_element_by_id("username")
    elem.send_keys("DZBS0453")
    elem2 = browser.find_element_by_id("password")
    time.sleep(1)
    elem2.send_keys("Soge2016*" + Keys.RETURN)
    time.sleep(4)
    import pythoncom
    pythoncom.CoInitialize()
    try:
        browser.find_element_by_id('btnContinue').click()
    except:
        pass
    time.sleep(4)
    shell = win32com.client.Dispatch("WScript.Shell")
    time.sleep(1)
    shell.SendKeys("{TAB}", 0)
    time.sleep(1)
    shell.SendKeys("{TAB}", 0)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    time.sleep(5)
    shell.SendKeys("{ENTER}", 0)

    pass


def ejecutar_ipon():
    browser = set_up_browser()
    login(browser)

# ejecutar_ipon()