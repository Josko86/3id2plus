from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import time

# set up browser
binary = FirefoxBinary(r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe')
browser = webdriver.Firefox(firefox_binary=binary)
browser.get('https://dro.orange-business.com/authentification?target=https://espaceclient.orange-business.com/group'
           '/divop/home?codeContexte=ece_divop&TYPE=33554433&REALMOID=06-00006a03-1ec3-1184-b5ad-5e0e0a63d064&'
           'GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-hCJfMsHRC8Nvudq1lQDbznywak%2fg%2bYsE6nklHqOsEk8XmYpdaqDy'
           'ezDHzkpWx6GU&TARGET=-SM-https%3a%2f%2fespaceclient%2eorange--business%2ecom%2fgroup%2fdivop%2fhome')

# login
elem = browser.find_element_by_id("username")
elem.send_keys("michael.yniesta")
elem2 = browser.find_element_by_id("password")
time.sleep(1)
elem2.send_keys("Scopelec92!" + Keys.RETURN)
time.sleep(4)

# boutique operations
browser.get('https://espaceclient.orange-business.com/group/divop/boutique-operateurs')
time.sleep(3)
main_window_handle = browser.current_window_handle

browser.switch_to_frame(browser.find_element_by_id('ece_iframe'))
a = browser.find_element_by_css_selector('.sfci_box > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > '
                                         'table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2)'
                                         ' > select:nth-child(1) > option:nth-child(7)')
a.click()
time.sleep(1)
# Elegir operation, esperamos para que se carguen las opciones
b = browser.find_element_by_css_selector('.sfci_box > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > '
                                         'table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(2) '
                                         '> select:nth-child(1) > option:nth-child(8)')
b.click()
c = browser.find_element_by_css_selector('.sfci_box > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) '
                                         '> table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(2)'
                                         ' > select:nth-child(1) > option:nth-child(2)')
c.click()
time.sleep(1)
browser.find_element_by_css_selector('a.sfci_blackb:nth-child(5)').click() #pulsamos en 'deposer a partir d'une...
# se abre una nueva ventana y hay que elegir el que corresponda con formato = SC1_” + nombre ciudad + “SPL” o “CPL” O “STR
time.sleep(1)
signin_window_handle = None
while not signin_window_handle:
    for handle in browser.window_handles:
        if handle != main_window_handle:
            signin_window_handle = handle
            break
browser.switch_to.window(signin_window_handle)
client = 'SC1'
city = 'ST_GERMAIN_LAXIS'
dos_type = 'SPL'
time.sleep(2)

row_text = browser.find_element_by_xpath("//*[contains(text(), 'SC1_ST_GERMAIN_LAXIS_SPL')]")
row_parent = row_text.find_element_by_xpath('..')
clickable_button = row_parent.find_element_by_xpath('.//input')
clickable_button.click()

time.sleep(2)
valider_button = browser.find_element_by_css_selector('a.sfci_blackb:nth-child(2)') # click on button valider
valider_button.click()
browser.switch_to.window(main_window_handle)
time.sleep(1)

# Volvemos a la pagina anterior y aparece un formulario en el que hay que rellenar algunos campos
aval_or_amont = 'aval'
aval1_info = 'F12312313123'
aval2_info = 'f0000123000'
arquetas_postes = 'postes'
time.sleep(1)
commande = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_pm_td_input '
                                                '> select:nth-child(1)')
commande.click()
time.sleep(1)
if aval_or_amont == 'aval':
    #opción marcada por defecto. No modificar el select, solo los 2 campos siguientes
    # aval_choice = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_pm'
    #                                                    '_td_input > select:nth-child(1) > option:nth-child(3)')
    # aval_choice.click()
    aval1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_cde_aval_pm_td'
                                                 '_input > input:nth-child(1)')
    aval1.clear()
    time.sleep(1)
    aval1.send_keys(aval1_info)
    aval2 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_ipe_td_input '
                                                 '> input:nth-child(1)')
    aval2.clear()
    aval2.send_keys(aval2_info)
else:
    amont_choice = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_'
                                                        'pm_td_input > select:nth-child(1) > option:nth-child(2)')
    amont_choice.click()

time.sleep(2)

if arquetas_postes == 'postes':
    select_postes = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:cont_com\/gcblo\:cde_'
                                                           'concerne > option:nth-child(4)')
    select_postes.click()




b = 2