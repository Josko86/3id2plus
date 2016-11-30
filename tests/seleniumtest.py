from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import time
import openpyxl


# Carga de todos los datos necesarios del excel



def set_up_browser():
    # set up browser
    binary = FirefoxBinary(r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe')
    browser = webdriver.Firefox(firefox_binary=binary)
    browser.get('https://dro.orange-business.com/authentification?target=https://espaceclient.orange-business.com/group'
               '/divop/home?codeContexte=ece_divop&TYPE=33554433&REALMOID=06-00006a03-1ec3-1184-b5ad-5e0e0a63d064&'
               'GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-hCJfMsHRC8Nvudq1lQDbznywak%2fg%2bYsE6nklHqOsEk8XmYpdaqDy'
               'ezDHzkpWx6GU&TARGET=-SM-https%3a%2f%2fespaceclient%2eorange--business%2ecom%2fgroup%2fdivop%2fhome')
    return browser

def login(browser):
    # login
    elem = browser.find_element_by_id("username")
    elem.send_keys("michael.yniesta")
    elem2 = browser.find_element_by_id("password")
    time.sleep(1)
    elem2.send_keys("Scopelec92!" + Keys.RETURN)
    time.sleep(4)


def boutique_operations(browser):
    # boutique operations
    dos_type = 'STR'
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
    if dos_type == 'SPL':
        b = browser.find_element_by_css_selector('.sfci_box > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > '
                                                 'table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(2) '
                                                 '> select:nth-child(1) > option:nth-child(8)')
    elif dos_type == 'CPL':
        b = browser.find_element_by_css_selector('.sfci_box > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > '
                                                 'table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(2)'
                                                 ' > select:nth-child(1) > option:nth-child(7)')
    elif dos_type == 'STR':
        b = browser.find_element_by_css_selector('.sfci_box > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1)'
                                                 ' > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > '
                                                 'td:nth-child(2) > select:nth-child(1) > option:nth-child(4)')

    b.click()
    c = browser.find_element_by_css_selector('.sfci_box > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) '
                                             '> table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(2)'
                                             ' > select:nth-child(1) > option:nth-child(2)')
    c.click()
    time.sleep(1)
    browser.find_element_by_css_selector('a.sfci_blackb:nth-child(5)').click() #pulsamos en 'deposer a partir d'une...
    # se abre una nueva ventana y hay que elegir el que corresponda con formato = cliente_” + nombre ciudad_ + “SPL” o “CPL” O “STR
    time.sleep(1)
    signin_window_handle = None
    while not signin_window_handle:
        for handle in browser.window_handles:
            if handle != main_window_handle:
                signin_window_handle = handle
                break
    browser.switch_to.window(signin_window_handle)
    client = 'SC1'
    city = 'COLOMBES'
    time.sleep(5)

    row_text = browser.find_element_by_xpath("//*[contains(text(), '" + client + '_' + city + '_' + dos_type + "')]")
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
    aval_primera = False
    aval1_info = 'F12312313123'
    aval2_info = 'f0000123000'
    arquetas_postes = 'postes'
    fecha_ini = '23/11/2017'
    fecha_fin = '21/12/2017'
    calles = ['rue del percebe', 'calle street', 'callejon hammer']
    time.sleep(1)
    commande = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_pm_td_input '
                                                    '> select:nth-child(1)')
    commande.click()
    time.sleep(1)

    # Para dosier Simple
    if dos_type == 'SPL':
        if aval_or_amont == 'aval':
            #opción marcada por defecto. No modificar el select, solo los 2 campos siguientes

            aval1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_cde_aval_pm_td'
                                                         '_input > input:nth-child(1)')
            aval1.clear()
            time.sleep(1)
            aval1.send_keys(aval1_info)
            aval2 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_ipe_td_input '
                                                         '> input:nth-child(1)')
            aval2.clear()
            aval2.send_keys(aval2_info)

        elif aval_or_amont == 'amont':
            amont_choice = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_'
                                                                'pm_td_input > select:nth-child(1) > option:nth-child(2)')
            amont_choice.click()

        time.sleep(2)

        # Si solo arquetas se deja por defecto GC, si no se pone GC et apus aeris
        if arquetas_postes == 'postes':
            select_postes = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:cont_com\/gcblo\:cde_'
                                                                   'concerne > option:nth-child(4)')
            select_postes.click()

        time.sleep(1)
        # Fechas de simple
        form_fecha_ini = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contenu_decl_trvx\/gcblo\:'
                                                              'date_debut_trvx')
        form_fecha_fin = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contenu_decl_trvx\/gcblo\:'
                                                              'date_fin_trvx_td_input > input:nth-child(1)')
        time.sleep(1)
        form_fecha_ini.clear()
        form_fecha_ini.send_keys(fecha_ini)
        time.sleep(1)
        form_fecha_fin.clear()
        form_fecha_fin.send_keys(fecha_fin)

        time.sleep(1)
    #     Annadimos las calles que pasan por el recorrido
        for i in range(len(calles)):
            calle_form = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contenu_decl_trvx\/gcblo\:'
                                                              'arteres_princ\['+str(i+1)+'\]_td_input > input:nth-child(1)')
            time.sleep(1)
            calle_form.clear()
            calle_form.send_keys(calles[i])
            if i < len(calles) - 1:
                time.sleep(1)
                add_button = browser.find_element_by_css_selector(
                    '#\/com\:commande\/gcblo\:contenu_decl_trvx\/gcblo\:arteres_princ\[1\]_td_label1 > a:nth-child(1) > '
                    'img:nth-child(1)')
                add_button.click()

    # Si el dosier es complejo o estructurante
    # a = zone de commande
    # b = IPE de PM
    # c = type de commande
    # d = Num el
    # e = cable mixte
    # f = FCI del primer
    if dos_type == 'CPL' or dos_type == 'STR':
        if aval_or_amont == 'aval' and not aval_primera:
            ipe = 'ipedelpm'
            fci_anterior = 'F21351351341'

            time.sleep(1)
            b = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_ipe_td_input > '
                                                     'input:nth-child(1)')
            b.clear()
            b.send_keys(ipe)
            f = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_cde_aval_pm_td_input > '
                                                     'input:nth-child(1)')
            f.clear()
            time.sleep(1)
            f.send_keys(fci_anterior)

        if aval_or_amont == 'aval' and aval_primera:
            ipe = 'ipedelpm'
            num_el = '5'

            time.sleep(1)
            b = browser.find_element_by_css_selector(
                '#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_ipe_td_input > input:nth-child(1)')
            b.clear()
            b.send_keys(ipe)
            c = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:type_cde_td_input > '
                                                     'select:nth-child(1)')
            c.click()
            time.sleep(1)
            c1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:type_cde_td_input'
                                                      ' > select:nth-child(1) > option:nth-child(2)')
            c1.click()
            d = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:taille_pm_td_input'
                                                     ' > input:nth-child(1)')
            d.send_keys(num_el)
            time.sleep(1)
            e = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:cab_mixte_td_input'
                                                     ' > select:nth-child(1)')
            e.click()
            time.sleep(1)
            e1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:cab_mixte_td_input'
                                                      ' > select:nth-child(1) > option:nth-child(3)')
            e1.click()
            f = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_cde_aval_pm_td_input'
                                                     ' > input:nth-child(1)')
            f.clear()
            time.sleep(1)

        if aval_or_amont == 'amont':
            a = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_pm_td_input'
                                                     ' > select:nth-child(1)')
            a.click()
            time.sleep(1)
            a1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_pm_td_input'
                                                      ' > select:nth-child(1) > option:nth-child(3)')
            a1.click()
            b = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_ipe_td_input'
                                                     ' > input:nth-child(1)')
            b.clear()
            c = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:type_cde_td_input'
                                                     ' > select:nth-child(1)')
            c.click()
            time.sleep(1)
            c1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:type_cde_td_input'
                                                      ' > select:nth-child(1) > option:nth-child(1)')
            c1.click()


        # Fechas de cpl y str
        time.sleep(1)
        # Si solo arquetas se deja por defecto GC, si no se pone GC et apus aeris
        if arquetas_postes == 'postes':
            select_postes = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:cont_com\/gcblo\:cde_'
                                                                 'concerne > option:nth-child(4)')
            select_postes.click()

        time.sleep(1)
        if dos_type == 'CPL':
            diferencial_tipo = 'contenu_declaration_travaux'
        elif dos_type == 'STR':
            diferencial_tipo = 'declaration_trvx'

        form_fecha_ini = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:' + diferencial_tipo + '\/gcblo\:date_debut_trvx')
        form_fecha_fin = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:' + diferencial_tipo + '\/gcblo\:date_fin_trvx_td_input > input:nth-child(1)')
        time.sleep(1)
        form_fecha_ini.clear()
        form_fecha_ini.send_keys(fecha_ini)
        time.sleep(1)
        form_fecha_fin.clear()
        form_fecha_fin.send_keys(fecha_fin)

        time.sleep(1)
        # Annadimos las calles que pasan por el recorrido
        for i in range(len(calles)):
            calle_form = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:' + diferencial_tipo + '\/gcblo\:arteres_princ\[' + str(i + 1) + '\]_td_input > input:nth-child(1)')
            time.sleep(1)
            calle_form.clear()
            calle_form.send_keys(calles[i])
            if i < len(calles) - 1:
                time.sleep(1)
                add_button = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:' + diferencial_tipo + '\/gcblo\:arteres_princ\[1\]_td_label1 > a:nth-child(1) > img:nth-child(1)')
                add_button.click()



# COMIENZA EL PROCESO
browser = set_up_browser()
login(browser)
boutique_operations(browser)
