# -*- coding: utf-8 -*-
import signal

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import time
import logging
import openpyxl
# import win32com.client as win32
from pyautocad import Autocad, APoint
import ezdxf
import shutil, os


# Funciones utils
def esAval(row, hp):
    global aval
    if hp.cell(row= row, column= 6).value == 'Aval PM':
        aval = True
    elif hp.cell(row= row, column= 6).value == 'Amont PM':
        aval = False
    return aval


def calculoTipo(row, hp):
    global tipo
    if hp.cell(row= row, column= 20).value == 'SIMPLE':
        tipo = 'SPL'
    elif hp.cell(row= row, column= 20).value == 'COMPLEXE':
        tipo = 'CPL'
    elif hp.cell(row= row, column= 20).value == 'STRUCTURANTE':
        tipo = 'STR'
    return tipo


def calculoFechas(tipo, col, hp):
    global date
    if tipo == 'SPL':
        date = hp.cell(row= 1, column=col).value
    elif tipo == 'CPL':
        date = hp.cell(row= 2, column=col).value
    elif tipo == 'STR':
        date = hp.cell(row=3, column=col).value
    date = date.strftime('%d/%m/%Y')
    return date


def calculoArquetas(row, hp):
    global arquetas
    if hp.cell(row=row, column=25).value != 0 and hp.cell(row=row, column=26).value == 0:
        arquetas = 'gc'
    elif hp.cell(row=row, column=25).value != 0 and hp.cell(row=row, column=26).value != 0:
        arquetas = 'gt+p'
    elif hp.cell(row=row, column=25).value == 0 and hp.cell(row=row, column=26).value != 0:
        arquetas = 'p'
    return arquetas


def calculoNumel(row,hp):
    global num_el
    if hp.cell(row=row, column=713).value == None:
        num_el = '0'
    else:
        a = hp.cell(row=row, column=713).value
        num_el = str(a)
    return num_el


def calculoCalles(row,hp):
    global calles
    calles = []
    if hp.cell(row=row, column=7).value == 'D2 IMB':
        calles.append(hp.cell(row=row, column=12).value + ' ' + hp.cell(row=row, column=13).value)
    if hp.cell(row=row, column=14).value != None:
        otras_calles = hp.cell(row=row, column=14).value
        otras_calles = otras_calles.split(sep='/')
        for c in otras_calles:
            calles.append(c)
    return calles


def cargarDatosExcel():
# Carga de todos los datos necesarios del excel y los mete en un diccionario de dosieres
    print('Carga de todos los datos necesarios del excel y los mete en un diccionario de dosieres')
# doc = openpyxl.load_workbook('SuiviJRU.xlsx')
# doc.save('SuiviJRU.xlsx')

    if os.name == 'nt':
        doc = openpyxl.load_workbook('SuiviJRU.xlsx', data_only=True)
        # doc = openpyxl.load_workbook(r'Z:/03-PRODUCCION/0.CAFT/SC1/PRODUCCIÓN/Tab Suivi Prod/suivi prod general SC1 practica-JRU.xlsx', data_only=True)
    else:
        doc = openpyxl.load_workbook(r'/home/ubuntu/3id2plus/SuiviJRU.xlsx', data_only=True)
        # doc = openpyxl.load_workbook(r'/home/ubuntu/nas/NAS/03-PRODUCCION/0.CAFT/SC1/PRODUCCIÓN/Tab Suivi Prod/suivi prod general SC1 practica-JRU.xlsx', data_only=True)
    doc.get_sheet_names()
    hoja_principal = doc.get_sheet_by_name('Tab Suivi Prod')
    dosieres = dict()

    for row in hoja_principal.iter_rows(min_row=1, max_col=1, max_row=hoja_principal.max_row):
        for celda in row:
            #TODO if celda.value in dosieres_act:  -->  Para seleccionar los dosieres que se hacen
            if hoja_principal.cell(row=celda.row, column=20).value == 'SIMPLE' and \
                            hoja_principal.cell(row=celda.row, column=40).value == None and celda.row != 1565:
                dosier = {
                    'nombre': celda.value,
                    'ciudad': hoja_principal.cell(row= celda.row, column= 16).value,
                    'es_aval': esAval(celda.row, hoja_principal),
                    'tipo': calculoTipo(celda.row, hoja_principal),
                    'es_1ca': hoja_principal.cell(row= celda.row, column=19).value == '1er CA',
                    'IPE_PM': hoja_principal.cell(row= celda.row, column= 8).value,
                    'ref_1era_PM': hoja_principal.cell(row= celda.row, column= 15).value,
                    'num_EL': calculoNumel(celda.row, hoja_principal),
                    'cliente': 'SC1',
                    'solo_arquetas': calculoArquetas(celda.row, hoja_principal), # gc, gc+p, p
                    'calles': calculoCalles(celda.row, hoja_principal),
                    'ref_cli': hoja_principal.cell(row= celda.row, column=79).value,
                    'formulario': 'SC1_ST_GERMAIN_LAXIS_SPL', #TODO poner la columna que ponga el nombre del form
                    'row': celda.row
                }
                dosier['date_ini'] = calculoFechas(dosier['tipo'], 4, hoja_principal)
                dosier['date_fin'] = calculoFechas(dosier['tipo'], 5, hoja_principal)
                dosieres[dosier['nombre']] = dosier
    # doc.save('SuiviJRU.xlsx')
    return dosieres

############################################ EJEMPLO DOSIER ###########################
# dosier : {
#     'nombre': 'Les1303'
#     'cliente': 'SC1',
#     'tipo': 'SPL',
#     'ciudad': 'COLOMBES',
#     'es_aval': True,
#     'es_1ca': True,
#     'solo_arquetas': True
#     'IPE_PM': 'FI-63463-12312',
#     'ref_1era_PM': 'F1241231233',
#     'num_el': '34',
#     'date_ini': '05/12/2016',
#     'date_fin': '02/01/2016',
#     'calles': ['rue del percebe', 'calle street'],
#     'ref_cli': 'SC1_IMB_93045_C_00X1',
#     'fci': 'F234124123'
#     'row': 345
# }
######################################################################################

def set_up_browser():
    # set up browser
    print('set up browser')
    # Windows
    if os.name == 'nt':
        ############################# FIREFOX ###########################################
        binary = FirefoxBinary(r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe')
        browser = webdriver.Firefox(firefox_binary=binary)

        ############################ PHANTOMJS ##########################################
        # path = r'C:\Users\josko\PycharmProjects\josko\scripts\phantomjs-2.1.1-windows\bin\phantomjs.exe'
        # browser = webdriver.PhantomJS(executable_path=path)
    # Linux
    else:
        browser = webdriver.PhantomJS()



    browser.get('https://dro.orange-business.com/authentification?target=https://espaceclient.orange-business.com/group'
               '/divop/home?codeContexte=ece_divop&TYPE=33554433&REALMOID=06-00006a03-1ec3-1184-b5ad-5e0e0a63d064&'
               'GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-hCJfMsHRC8Nvudq1lQDbznywak%2fg%2bYsE6nklHqOsEk8XmYpdaqDy'
               'ezDHzkpWx6GU&TARGET=-SM-https%3a%2f%2fespaceclient%2eorange--business%2ecom%2fgroup%2fdivop%2fhome')
    return browser

def login(browser):
    print('login')
    elem = browser.find_element_by_id("username")
    elem.send_keys("michael.yniesta")
    elem2 = browser.find_element_by_id("password")
    time.sleep(1)
    elem2.send_keys("Scopelec92!" + Keys.RETURN)
    time.sleep(4)


def boutique_operations(browser, d):
    print('boutique operations')
    dos_type = d['tipo']
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
    # se abre una nueva ventana y hay que elegir el formulario que corresponda
    time.sleep(1)
    # signin_window_handle = None
    # while not signin_window_handle:
    #     for handle in browser.window_handles:
    #         if handle != main_window_handle:
    #             signin_window_handle = handle
    #             break
    time.sleep(5)
    signin_window_handle = browser.window_handles[1]
    browser.switch_to.window(signin_window_handle)
    formulario = d['formulario']
    time.sleep(5)

    row_text = browser.find_element_by_xpath("//*[contains(text(), '" + formulario + "')]")
    row_parent = row_text.find_element_by_xpath('..')
    clickable_button = row_parent.find_element_by_xpath('.//input')
    clickable_button.click()

    time.sleep(2)
    valider_button = browser.find_element_by_css_selector('a.sfci_blackb:nth-child(2)') # click on button valider
    valider_button.click()
    time.sleep(2)
    browser.switch_to.window(main_window_handle)
    time.sleep(1)

    # Volvemos a la pagina anterior y aparece un formulario en el que hay que rellenar algunos campos
    es_aval = d['es_aval']
    aval_primera = d['es_1ca']
    ref_1era_PM = d['ref_1era_PM']
    IPE_PM = d['IPE_PM']
    solo_arquetas = d['solo_arquetas']
    fecha_ini = d['date_ini']
    fecha_fin = d['date_fin']
    calles = d['calles']
    # Si esta en phantomjs tiene que entrar en el iframe
    if browser.name == 'phantomjs':
        browser.switch_to_frame('ece_iframe')
    time.sleep(1)
    commande = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_pm_td_input '
                                                    '> select:nth-child(1)')
    commande.click()
    time.sleep(1)

    # Para dosier Simple
    if dos_type == 'SPL':
        aval1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_cde_aval_pm_td'
                                                     '_input > input:nth-child(1)')
        aval1.clear()
        time.sleep(1)
        aval2 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_ipe_td_input '
                                                     '> input:nth-child(1)')
        aval2.clear()
        time.sleep(1)
        if es_aval:
            #opción marcada por defecto. No modificar el select, solo los 2 campos siguientes
            aval_choice = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde'
                                                               '_pm_td_input > select:nth-child(1) > option:nth-child(3)')
            aval_choice.click()
            aval1.send_keys(ref_1era_PM)
            aval2.send_keys(IPE_PM)

        elif not es_aval:
            time.sleep(1)
            amont_choice = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_'
                                                                'pm_td_input > select:nth-child(1) > option:nth-child(2)')
            amont_choice.click()
            browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_'
                                                 'pm_td_input > select:nth-child(1) > option:nth-child(2)').click()

        time.sleep(2)

        # Si solo arquetas se deja por defecto GC, si arquetas + poteaux se pone GC et apus aeris
        select_postes = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:cont_com\/gcblo\:cde_'
                                                             'concerne > option:nth-child(2)')
        if solo_arquetas == 'gc+p':
            select_postes = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:cont_com\/gcblo\:cde_'
                                                                 'concerne > option:nth-child(4)')
        elif solo_arquetas == 'p':
            select_postes = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:cont_com\/gcblo\:cde_'
                                                                 'concerne > option:nth-child(3)')
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
    # af = zone de commande
    # bf = IPE de PM
    # cf = type de commande
    # df = Num el
    # ef = cable mixte
    # ff = FCI del primer
    if dos_type == 'CPL' or dos_type == 'STR':
        bf = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_ipe_td_input > '
                                                  'input:nth-child(1)')
        bf.clear()
        time.sleep(1)
        df = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:taille_pm_td_input'
                                                  ' > input:nth-child(1)')
        df.clear()
        time.sleep(1)
        ff = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:ref_cde_aval_pm_td_'
                                                  'input > input:nth-child(1)')
        ff.clear()
        time.sleep(1)

        if es_aval and not aval_primera:
            af = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_pm_'
                                                      'td_input > select:nth-child(1) > option:nth-child(2)')
            af.click()
            time.sleep(1)
            bf.send_keys(IPE_PM)
            time.sleep(1)
            cf = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:type_cde_td_input'
                                                      ' > select:nth-child(1) > option:nth-child(3)')
            cf.click()
            time.sleep(1)
            ff.send_keys(ref_1era_PM)

        if es_aval and aval_primera:
            num_el = d['num_EL']
            af = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:'
                                                      'zone_cde_pm_td_input > select:nth-child(1) > option:nth-child(2)')
            af.click()
            time.sleep(1)
            bf.send_keys(IPE_PM)
            cf = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:type_cde_td_input > '
                                                      'select:nth-child(1)')
            cf.click()
            time.sleep(1)
            c1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:type_cde_td_input'
                                                      ' > select:nth-child(1) > option:nth-child(2)')
            c1.click()
            df.send_keys(num_el)
            time.sleep(1)
            ef = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:cab_mixte_td_input'
                                                      ' > select:nth-child(1)')
            ef.click()
            time.sleep(1)
            e1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:cab_mixte_td_input'
                                                      ' > select:nth-child(1) > option:nth-child(3)')
            e1.click()
            time.sleep(1)

        if not es_aval:
            time.sleep(1)
            af = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_pm_td_'
                                                      'input > select:nth-child(1) > option:nth-child(3)')
            af.click()
            time.sleep(1)
            a1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:zone_cde_pm_td_input'
                                                      ' > select:nth-child(1) > option:nth-child(3)')
            a1.click()
            cf = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:type_cde_td_input'
                                                      ' > select:nth-child(1)')
            cf.click()
            time.sleep(1)
            c1 = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:contexte_com\/gcblo\:type_cde_td_input'
                                                      ' > select:nth-child(1) > option:nth-child(1)')
            c1.click()

        time.sleep(1)
        # Si solo arquetas se deja por defecto GC, si arquetas + poteaux se pone GC et apus aeris
        select_postes = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:cont_com\/gcblo\:cde_'
                                                             'concerne > option:nth-child(2)')
        if solo_arquetas == 'gc+p':
            select_postes = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:cont_com\/gcblo\:cde_'
                                                                 'concerne > option:nth-child(4)')
        elif solo_arquetas == 'p':
            select_postes = browser.find_element_by_css_selector('#\/com\:commande\/gcblo\:cont_com\/gcblo\:cde_'
                                                                 'concerne > option:nth-child(3)')
        select_postes.click()

        # Fechas de cpl y str
        time.sleep(1)
        diferencial_tipo = 'declaration_trvx'
        if dos_type == 'CPL':
            diferencial_tipo = 'contenu_declaration_travaux'

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

                # TODO recuperar el numero fci
    fci = 'F4214120316'
    d['fci'] = fci
    b = 1


def tsp_operations_1(dosier, ws):
# operaciones despues de obtener FCI en el tsp
    ws.Cells(dosier['row'], 40).Value = dosier['fci']
    fecha_fci = dosier['fci'][-4:-2] + '-' + dosier['fci'][-6:-4] + '-' + dosier['fci'][-2:]
    ws.Cells(dosier['row'], 41).Value = fecha_fci
    ws.Cells(dosier['row'], 42).Value = 'DEP'
    ws.Cells(dosier['row'], 65).Value = 'v1'
    ws.Cells(dosier['row'], 66).Value = 'Attente Input TFX'


#  TODO
def mover_ficheros(d):
    ruta = os.getcwd() + os.sep
    origen = ruta + 'movido'
    destino = 'C:\\Users\\josko\\PycharmProjects\\josko\\mover\\'
    if os.path.exists(origen):
        ruta = shutil.move(origen, destino)
        print('El directorio ha sido movido a', ruta)
    else:
        print('El directorio origen no existe')
    pass



###################################### COMIENZA EL PROCESO ##################################################

def obtenerFCI():
    logging.basicConfig(filename='webop.log',level=logging.INFO,
                        format='%(asctime)s %(levelname)s: %(message)s', datefmt='%d/%m/%Y %I:%M:%S %p')


    # dosieres de prueba para no tener que cargar el excel continuamente
    # dosieres = {
    #             'Sus1314': {'es_aval': True, 'IPE_PM': 'FI-92073-0023', 'ref_1era_PM': 'A DEPOSER', 'date_fin': '18/04/2017', 'calles': ['rue del percebe', 'calle street', 'callejon hammer'], 'ref_cli': '', 'solo_arquetas': True, 'ciudad': 'SURESNES', 'date_ini': '28/12/2016', 'nombre': 'Sus1314', 'es_1ca': True, 'cliente': 'SC1', 'tipo': 'STR', 'num_EL': 38},
    #             'Sus1136': {'es_aval': True, 'IPE_PM': 'FI-92073-0017', 'ref_1era_PM': 'F34837031016', 'date_fin': '18/01/2017', 'calles': ['rue del percebe', 'calle street', 'callejon hammer'], 'ref_cli': '', 'solo_arquetas': True, 'ciudad': 'SURESNES', 'date_ini': '14/12/2016', 'nombre': 'Sus1136', 'es_1ca': False, 'cliente': 'SC1', 'tipo': 'CPL', 'num_EL': 18},
    #             'Sus1269': {'es_aval': False, 'IPE_PM': None, 'ref_1era_PM': None, 'date_fin': '18/01/2017', 'calles': ['rue del percebe', 'calle street', 'callejon hammer'], 'ref_cli': '', 'solo_arquetas': True, 'ciudad': 'SURESNES', 'date_ini': '14/12/2016', 'nombre': 'Sus1269', 'es_1ca': False, 'cliente': 'SC1', 'tipo': 'CPL', 'num_EL': 0},
    #             'Moe993': {'es_aval': False, 'IPE_PM': None, 'ref_1era_PM': None, 'date_fin': '18/04/2017', 'calles': ['rue del percebe', 'calle street', 'callejon hammer'], 'ref_cli': '', 'solo_arquetas': True, 'ciudad': 'MONTEREAU-FAULT-YONNE', 'date_ini': '28/12/2016', 'nombre': 'Moe993', 'es_1ca': False, 'cliente': 'SC1', 'tipo': 'STR', 'num_EL': 0},
    #             'Cly1345': {'es_aval': False, 'IPE_PM': None, 'ref_1era_PM': None, 'date_fin': '03/01/2017', 'calles': ['rue del percebe', 'calle street', 'callejon hammer'], 'ref_cli': '', 'solo_arquetas': True, 'ciudad': 'CLICHY', 'date_ini': '06/12/2016', 'nombre': 'Cly1345', 'es_1ca': False, 'cliente': 'SC1', 'tipo': 'SPL', 'num_EL': 17},
    #             'Sus1230': {'es_aval': True, 'IPE_PM': 'FI-92073-001E', 'ref_1era_PM': 'F28968041116', 'date_fin': '18/01/2017', 'calles': ['rue del percebe', 'calle street', 'callejon hammer'], 'ref_cli': '', 'solo_arquetas': True, 'ciudad': 'SURESNES', 'date_ini': '14/12/2016', 'nombre': 'Sus1230', 'es_1ca': True, 'cliente': 'SC1', 'tipo': 'CPL', 'num_EL': 10}
    #
    # }
    dosieres = {
        'Nee1590': {'formulario': 'SC1_ST_GERMAIN_LAXIS_SPL', 'num_EL': None, 'solo_arquetas': 'gc', 'IPE_PM': None, 'es_aval': False, 'cliente': 'SC1', 'tipo': 'SPL', 'ref_1era_PM': None, 'row': 132, 'calles': ['RUE  BAILLY'], 'date_fin': '02/03/2017', 'ciudad': 'NEUILLY SUR SEINE', 'nombre': 'Nee1590', 'es_1ca': False, 'ref_cli': 'SC1_IMB_92051_C_3178', 'date_ini': '02/02/2017'}
                }

    #  FUNCIONA PARA ACCEDER AL NAS DESDE MI ORDENADOR
    # for cosa in os.listdir('Z:/03-PRODUCCION/0.CAFT/SC1/PRODUCCIÓN/Tab Suivi Prod'):
    #     b = cosa
    #     a = 2

    #TODO probar autocad, autoit, zip files...

    # acad = Autocad.('Fxxxxxxxxxxx_92078_xxxxxx.dxf')
    # acad.prompt("Hello, Autocad from Python\n")
    # print(acad.doc.Name)
    #
    # pass
    # b= 1

    try:
        dosieres = cargarDatosExcel()
    except Exception as ex:
        logging.error('No han podido cargarse los datos del excel porque: %s', ex.msg)
    result = {}

    import win32com.client as win32
    import pythoncom
    pythoncom.CoInitialize()

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    if os.name == 'nt':
        wb = excel.Workbooks.Open(r'C:\Users\josko\PycharmProjects\josko\SuiviJRU.xlsx')
    else:
        wb = excel.Workbooks.Open(r'\home\ubuntu\3id2plus\SuiviJRU.xlsx')
    ws = wb.Worksheets('Tab Suivi Prod')

#TODO cambiar a keys de dosieres antes estaba en dosieres_act
    for d in dosieres:
        try:
            browser = set_up_browser()
            login(browser)
            boutique_operations(browser, dosieres[d])
            time.sleep(4)
            tsp_operations_1(dosieres[d], ws)
            # mover_ficheros(dosieres[d])

        except Exception as ex:
            logging.error('%s No ha podido completarse por: %s', dosieres[d]['nombre'], ex.msg)
            result[dosieres[d]['nombre']] = dosieres[d]['nombre'] + ' No ha podido completarse por: ' + ex.msg

        else:
            logging.info('%s --> Se ha procesado correctamente: ', dosieres[d]['nombre'])
            result[dosieres[d]['nombre']] = dosieres[d]['nombre'] + '  --> Se ha procesado correctamente: '

        finally:
            if os.name != 'nt':
                browser.service.process.send_signal(signal.SIGTERM)
            browser.quit()
            time.sleep(5)

    print('Salvando excel')
    wb.SaveAs(r'C:\Users\josko\PycharmProjects\josko\SuiviJRU2.xlsx')
    excel.Application.Quit()
    return result


