import os

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import time
import win32com.client
import win32api, win32con
import math
import pythoncom

pythoncom.CoInitialize()
shell = win32com.client.Dispatch("WScript.Shell")
INTERVAL = 25

def win32_click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


def set_up_browser():
    # set up browser
    # Windows
    if os.name == 'nt':
        ############################# FIREFOX ###########################################
        # binary = FirefoxBinary(r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe')
        # profile = webdriver.FirefoxProfile(r'C:\Users\josko\AppData\Roaming\Mozilla\Firefox\Profiles\b6lwaicj.default')
        # capabilities = webdriver.DesiredCapabilities().FIREFOX
        # capabilities['acceptSslCerts'] = True
        # profile.accept_untrusted_certs = True
        # browser = webdriver.Firefox(firefox_binary=binary, firefox_profile=profile, capabilities=capabilities)
        # browser.capabilities['acceptSslCerts'] = True
        ################################## IE ####################################################
        browser = webdriver.Ie(r'C:\Users\josko\PycharmProjects\josko\scripts\IEDriverServer.exe')
    browser.get('https://lyon.metagate.orange.com/dana/home/index.cgi')
    return browser


def login(browser):

    elem = browser.find_element_by_id("username")
    elem.send_keys("R")
    shell.SendKeys('{DOWN}', 0)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    elem2 = browser.find_element_by_id("password")
    # elem2.send_keys("Soge2016*")
    time.sleep(1)
    elem2.send_keys(Keys.RETURN)
    time.sleep(2)
    try:
        browser.find_element_by_id('btnContinue').click()
    except:
        pass
    time.sleep(10)
    ipon_link = browser.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/center/table/tbody/tr/td/div[1]/div/table/tbody/tr/td/table/tbody/tr/td/div[1]/table[5]/tbody/tr/td[1]/table/tbody/tr/td[2]/a/b')
    ipon_link.click()
    time.sleep(4)
    main_window = browser.window_handles[0]
    second_window = browser.window_handles[1]
    browser.switch_to_window(second_window)
    time.sleep(1)
    browser.close()
    time.sleep(1)
    browser.switch_to_window(main_window)
    time.sleep(1)
    browser.get('http://ipon.sso.francetelecom.fr/NGI/GassiAccess.jsp')
    time.sleep(1)
    user_form = browser.find_element_by_id('user')
    user_form.send_keys('R')
    time.sleep(1)
    shell.SendKeys('{DOWN}', 0)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    time.sleep(1)
    shell.SendKeys("{TAB}", 0)
    time.sleep(1)
    shell.SendKeys("{TAB}", 0)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    time.sleep(3)
    browser.get('http://ipon.sso.francetelecom.fr/NGI/GassiAccess.jsp')
    time.sleep(4)


def crear_proyecto_ipon(browser, nra):

    # Pulsar mon bureau
    browser.get('http://ipon.sso.francetelecom.fr/desktop.jsp')
    time.sleep(3)
    nouveau_project = browser.find_element_by_xpath('/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr/td[2]/a')
    nouveau_project.click()
    time.sleep(3)
    c5 = 'prue'
    c9 = 'josko'
    nom = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[1]/td[2]/input')
    nom.clear()
    nom.send_keys('_'.join([c5, c9]))
    code_secteur = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[5]/td[2]/input')
    code_secteur.clear()
    code_secteur.send_keys(nra)
    code_oeie = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[6]/td[2]/input')
    code_oeie.clear()
    code_oeie.send_keys('000000')
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[1]/tbody/tr/td[4]/a').click()
    time.sleep(2)


def estudio(browser, nra):

    browser.find_element_by_xpath('/html/body/div/div[1]/table/tbody/tr[1]/td/table/tbody/tr/td[3]').click()
    time.sleep(1)
    research_immueble = browser.find_element_by_xpath('/html/body/div/div[1]/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/ul/li/ul/li[14]/a/span')
    research_immueble.click()
    time.sleep(1)
    id_inmuble_form = browser.find_element_by_xpath(
        '/html/body/div[1]/div[1]/div/form[1]/div/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[1]/table/tbody/tr/td[3]/font/div/input')
    time.sleep(1)
    id_inmuble_form.send_keys("I")
    shell.SendKeys('{DOWN}', 0)  # Elegir el imnmueble
    time.sleep(1)
    shell.SendKeys('{DOWN}', 0)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    time.sleep(6)
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form[2]/table/tbody/tr/td/div[1]/table/tbody/tr/td[2]/a').click()
    time.sleep(3)
    """
    # modifier IMB
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[1]/tbody/tr/td[2]/a').click()
    time.sleep(3)
    # seleccionar OI
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[6]/td[2]/select').click()
    time.sleep(1)
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[6]/td[2]/select/option[2]').click()
    time.sleep(1)
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[12]/td[2]/select').click()
    time.sleep(1)
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[12]/td[2]/select/option[3]').click()
    time.sleep(1)
    # abrimos nueva ventana para seleccionar el NRA
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[3]/td[2]/a[1]').click()
    time.sleep(5)
    main_window = browser.current_window_handle
    signin_window_handle = browser.window_handles[1]
    browser.switch_to.window(signin_window_handle)
    time.sleep(1)
    frame = browser.find_element_by_xpath('/html/frameset/frame[1]')
    frame2 = browser.find_element_by_xpath('/html/frameset/frame[2]')
    browser.switch_to_frame(frame)
    time.sleep(1)
    row_text = browser.find_element_by_xpath("//*[contains(text(), 'NRA " + nra + "')]")
    row_parent = row_text.find_element_by_xpath('../..')
    clickable_button = row_parent.find_element_by_xpath('td[1]/input')
    clickable_button.click()
    time.sleep(2)
    browser.switch_to_default_content()
    browser.switch_to_frame(frame2)
    time.sleep(1)
    # click selectionner
    browser.find_element_by_xpath('/html/body/form/div[1]/div/a').click()
    time.sleep(1)
    browser.switch_to_window(main_window)
    time.sleep(1)
    # mettre a jours (save)
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[1]/tbody/tr/td[2]/a').click()
    time.sleep(2)
    """
    # consulter metre la jour
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[1]/tbody/tr/td[4]/a').click()
    time.sleep(3)
    # browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[10]/font/a[2]').click()
    # time.sleep(2)
    # browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/form/table/tbody/tr[3]/td/font/div[1]/a').click()
    # time.sleep(3)
    browser.set_window_position(2213, 274)
    browser.set_window_size(1700, 1100)
    shell.SendKeys("{F12}", 0)
    time.sleep(6)
    win32_click(2397, 757)
    time.sleep(1)
    # win32_click(2890, 556)
    # time.sleep(1)
    # for i in range(4):
    #     shell.SendKeys('{UP}', 0)
    #     time.sleep(1)
    # shell.SendKeys('{DOWN}', 0)
    # time.sleep(1)
    # shell.SendKeys("{ENTER}", 0)
    # time.sleep(1)
    # win32_click(2994, 555)
    # time.sleep(1)
    # for i in range(4):
    #     shell.SendKeys('{DOWN}', 0)
    #     time.sleep(1)
    # shell.SendKeys('{UP}', 0)
    # time.sleep(1)
    # shell.SendKeys("{ENTER}", 0)
    # time.sleep(2)
    # # click sauvegarder
    # browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/'
    #                               'table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[2]/font/a[2]').click()
    # time.sleep(3)
    # # click [no name]
    browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/tabl'
                                  'e[2]/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr/td[2]/div/a').click()
    time.sleep(1)
    # click nouvel escalier
    browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td/form/'
                                  'table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[5]/font/a[2]').click()
    time.sleep(2)
    browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/form/table/tbody/tr[5]/'
                                  'td/font/div[1]/a').click()
    time.sleep(3)
    # click nouveau niveau
    browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/'
                                  'table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[5]/font/a[2]').click()
    nombre_form = browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/font/input')
    nombre_form.clear()
    nombre_form.send_keys('3')
    apartirde_form = browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/form/table/tbody/tr[2]/td[2]/font/input')
    apartirde_form.clear()
    apartirde_form.send_keys('0')
    time.sleep(1)
    browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/form/table/tbody/tr[4]/td/font/div[1]/a').click()
    time.sleep(4)
    # type de lescalier
    win32api.SetCursorPos(3176,1084)
    time.sleep(1)
    win32_click(3176,1084)
    time.sleep(1)
    shell.SendKeys('{DOWN}', 0)
    time.sleep(1)
    shell.SendKeys('{DOWN}', 0)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    time.sleep(2)
    browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[2]/font/a[2]').click()
    time.sleep(3)
    # Conexiones de cada piso
    y= 547
    for i in range(3): #Numero de niveles
        win32_click(2700, y)
        time.sleep(2)
        shell.SendKeys('2', 0)
        time.sleep(1)
        shell.SendKeys("{ENTER}", 0)
        time.sleep(2)
        y = y + INTERVAL


    pass


def ejecutar_ipon(nra):
    browser = set_up_browser()
    login(browser)
    # crear_proyecto_ipon(browser, nra)
    estudio(browser, nra)


# ejecutar_ipon()
