import os

from pywin.tools import browser
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import time
import win32com.client
import win32api, win32con
import pythoncom
import openpyxl

pythoncom.CoInitialize()
shell = win32com.client.Dispatch("WScript.Shell")
INTERVAL = 25
CTH = 103
CTW = 2
WX = 2213
WY = 274

def win32_click(x, y):
    win32api.SetCursorPos((x, y))
    time.sleep(1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


def elem_but_pos(el, repeat_click= True) :
    loc = el.location
    size = el.size
    x = WX + loc['x'] + size['width'] + CTW
    y = WY + loc['y'] + size['height'] + CTH
    win32_click(x, y)
    if repeat_click:
        win32_click(x, y)


def select_pa(browser, pa_chambre, inmueble, frame2):
    browser.find_element_by_xpath(
        '/html/body/div[1]/div/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/ul/li/div/form/input[1]').click()
    time.sleep(1)
    rechercher_chambre = browser.find_element_by_xpath(
        '/html/body/div[1]/div[1]/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/ul/li/ul/li[16]/a/span')
    rechercher_chambre.click()
    time.sleep(2)
    code_chambre_form = browser.find_element_by_xpath(
        '/html/body/div[1]/div/div/form[1]/div/table/tbody/tr[2]/td[1]/table/tbody/tr[3]/td/div/table/tbody/tr/td[3]/font/span/input')
    code_chambre_form.send_keys(pa_chambre)
    time.sleep(1)
    browser.find_element_by_id('searchButton').click()
    time.sleep(4)
    insee = inmueble.split('/')[1]
    identifiant = pa_chambre + '/' + insee
    row_ch = browser.find_element_by_xpath("//*[contains(text(), '" + identifiant + "')]")
    row_parent = row_ch.find_element_by_xpath('..')
    ch = row_parent.find_element_by_xpath('td[1]').click()
    time.sleep(2)
    browser.switch_to_default_content()
    time.sleep(1)
    browser.switch_to_frame(frame2)
    time.sleep(1)
    browser.find_element_by_xpath('/html/body/form/div[1]/div/a').click()
    time.sleep(1)


def select_imb_con_pt(browser, inmueble, frame2):
    browser.find_element_by_xpath(
        '/html/body/div[1]/div/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/ul/li/div/form/input[1]').click()
    time.sleep(1)
    rechercher_inmueble = browser.find_element_by_xpath(
        '/html/body/div[1]/div[1]/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/ul/li/ul/li[14]/a/span')
    rechercher_inmueble.click()
    time.sleep(1)
    id_inmuble_form = browser.find_element_by_xpath(
        '/html/body/div[1]/div[1]/div/form[1]/div/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[1]/table/tbody/tr/td[3]/font/div/input')
    time.sleep(1)
    id_inmuble_form.send_keys('I')
    shell.SendKeys('{DOWN}', 0)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    time.sleep(8)
    row_imb = browser.find_element_by_xpath(
        '/html/body/div[1]/div/div/form[2]/table/tbody/tr/td/div[1]/table/tbody/tr/td[1]').click()
    time.sleep(2)
    browser.switch_to_default_content()
    time.sleep(1)
    browser.switch_to_frame(frame2)
    time.sleep(1)
    browser.find_element_by_xpath('/html/body/form/div[1]/div/a').click()
    time.sleep(1)


def select_pt_in_imb(browser, frame2, pt):
    row_pt = browser.find_element_by_xpath("//*[contains(text(), '" + pt + "')]")
    row_parent = row_pt.find_element_by_xpath('..')
    pt_selected = row_parent.find_element_by_xpath('td[1]').click()
    time.sleep(2)
    browser.switch_to_default_content()
    time.sleep(1)
    browser.switch_to_frame(frame2)
    time.sleep(1)
    browser.find_element_by_xpath('/html/body/form/div[1]/div/a').click()
    time.sleep(1)

def get_data():
    """
    proyecto {
        "id_IMB":
        "adress_IMB":
        "IMBs": {
            "IMB/51454/X/0314": {
                                "A": {
                                    "1" : {
                                        "tipo_material": '[6.i.aco]'
                                        "material": SI,
                                        "num_el" 2
                                        }
                                    "0": {
                                        "material": NO,
                                        "num_el" 1
                                        }
                                    },
                                "B": {
                                    "2" : {
                                        "material": NO,
                                        "num_el" 1
                                        }
                                    "1" : {
                                        "tipo_material": '[6.i.aco]'
                                        "material": SI,
                                        "num_el" 1
                                        }
                                    "0" : {
                                        "material": NO,
                                        "num_el": 3
                                        }
                                    }
                                },
            "IMB/51454/C/OMM0": {
                                "C": {
                                    "2": {
                                        "material": NO,
                                        "num_el": 1
                                        },
                                    "1": {
                                        "tipo_material": '[12.i13.3m]'
                                        "material" SI,
                                        "num_el": 2
                                        },
                                    "0": {
                                        "material" NO,
                                        "num_el": 2
                                        }
                                }
            }
        }





    Proyecto por ficha
        Cada PROYECTO puede tener varios IMB
            Cada IMB puede tener varios BATIMENT
                Cada BATIMENT tiene una COLONNE MONTANTE
                    Cada COLONNE MONTANTE puede tener varios NIVELES
                        Cada NIVEL tiene:
                            material: bool
                            num_el: int
    :return: proyecto
    """
    pythoncom.CoInitialize()
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(r'C:\Users\josko\PycharmProjects\josko\inmueble_prueba4.xls')
    excel.Visible = True
    ws_ic = wb.Worksheets('Infos clés')
    ws_pb = wb.Worksheets('PB')
    id_imb = ws_ic.Cells(5, 3).Value
    adress_imb = ws_ic.Cells(9, 3).Value
    project = dict()
    IMBs = dict()
    PBs = dict()
    BTIs = dict()
    cables_list = []
    #TODO cambiar cuando sea definitive
    project['nom_project'] = id_imb + '_' + 'josk'
    # project['nom_project'] = id_imb + '_' + adress_imb
    row = 6
    COL_CM = 2
    COL_MATERIAL = 7
    COL_EL = 9
    COL_IMB = 28
    COL_OBSERVATION = 17
    COL_CM_IMB = 20
    COL_TYPE_BTI = 21
    COL_NIVEL = 3
    COL_PB_NAME = 15
    row_imb = 7
    bti_ini = 1

    while True:   # Una vuelta por cada row de la columna 20
        cm_imb = ws_pb.Cells(row_imb, COL_CM_IMB)
        if cm_imb.Value is None:
            break
        if cm_imb.Value < 'A':  # Si es un número la columna
            BTIs[cm_imb.Value]['tipo'] = ws_pb.Cells(row_imb, COL_TYPE_BTI).Value
        imb = ws_pb.Cells(row_imb, COL_IMB)
        imb = str(imb)
        row_imb += 1
        cm = cm_imb
        while True:    # Una vuelta por cada row de la columna 2
            cm = ws_pb.Cells(row, COL_CM)
            if cm.Value != cm_imb.Value:
                break
            cm = str(cm)
            material = ws_pb.Cells(row, COL_MATERIAL)
            if material.Value: material = str(material)
            el = ws_pb.Cells(row, COL_EL)
            el = str(el)
            el = el[0:1]
            nivel = ws_pb.Cells(row, COL_NIVEL)
            nivel = str(nivel)
            if nivel == 'RC':
                nivel = '0'
            elif len(nivel) > 2:
                nivel = nivel[:-2]
            observation = ws_pb.Cells(row, COL_OBSERVATION).Value
            if observation is None: observation = ""
            bti = "BTI" in observation
            hay_material = False
            if type(material) == str or bti:
                hay_material = True
            pb_name = ws_pb.Cells(row, COL_PB_NAME).Value
            row += 1
            while True:  #  Una vuelta por cada elemento del nivel (row)
                if imb in IMBs:
                    if cm in IMBs[imb]:
                        if nivel in IMBs[imb][cm]:
                            IMBs[imb][cm][nivel]['pb_name'] = pb_name
                            IMBs[imb][cm][nivel]['material'] = hay_material
                            IMBs[imb][cm][nivel]['num_el'] = el
                            if pb_name in PBs:
                                PBs[pb_name]['num_el'] += int(el)
                                PBs[pb_name]['niveles'].append(nivel)
                            else:
                                PBs[pb_name] = {'num_el':int(el), 'colonne':cm, 'niveles':[nivel], 'inmueble': imb}

                            if hay_material:
                                if bti:
                                    IMBs[imb][cm][nivel]['observation'] = observation
                                    BTIs[str(bti_ini)] = {'nivel_is': nivel, 'colonne_is': cm, 'imb_is': imb, 'cms':[], 'observation':observation}
                                    bti_ini += 1
                                if type(material) == str:
                                    IMBs[imb][cm][nivel]['tipo_material'] = material
                                    PBs[pb_name]['tipo'] = material
                                    PBs[pb_name]['nivel'] = nivel
                                    PBs[pb_name]['observation'] = observation
                            break
                        else:
                            IMBs[imb][cm][nivel] = {}
                    else:
                        IMBs[imb][cm] = {'bti': ws_pb.Cells(row_imb - 1, 26).Value}
                        if IMBs[imb][cm]['bti'] == '<na>':
                            IMBs[imb][cm].pop('bti', None)
                else:
                    IMBs[imb] = {}
    for imb in IMBs:
        for cm in IMBs[imb]:
            for bti in BTIs:
                if bti in IMBs[imb][cm]['bti']:
                    BTIs[bti]['cms'].append(cm)

    project['inmuebles'] = IMBs
    project['pbs'] = PBs
    project['btis'] = BTIs
    wb.Close(False)
    wb = excel.Workbooks.Open(r'C:\Users\josko\PycharmProjects\josko\cablage4.xlsx')
    excel.Visible = True
    sheet = wb.Worksheets(1)
    canvas = sheet.Shapes
    for shp in canvas:
        box = shp.TextFrame2.TextRange.Characters.Text
        if 'CH ' in box and 'PA' in box:
            project['pa_chambre'] = box[-5:]
            project['pa_pt'] = box[-15:-9]

    # Recogemos los cables y los ponemos en en un diccionario con clave los extremos del cable Ej. PA-b
    for shp in canvas:
        box = shp.TextFrame2.TextRange.Characters.Text
        if 'TR ' in box:
            cables_list.append(box)
    cables = {}
    for c in cables_list:
        a = c.split(sep='\n')
        nombre = a[0][3:]
        ini = nombre.split(sep='-')[0]
        fin = nombre.split(sep='-')[1]
        b = a[1].split(sep=' ')
        metros = b[0]
        num_fo = b[3]
        cables[nombre] = {'metros': metros, 'num_fo': num_fo, 'ini': ini, 'fin': fin}

    project['cables'] = cables
    wb.Close(False)
    excel.Application.Quit()
    return project


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
    browser.set_window_position(WX, WY)
    browser.set_window_size(1700, 1100)
    time.sleep(1)
    return browser


def login(browser):

    elem = browser.find_element_by_id("username")
    elem.send_keys("L")
    shell.SendKeys('{DOWN}', 1)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    elem2 = browser.find_element_by_id("password")
    # elem2.send_keys("Soge2017*")
    time.sleep(1)
    elem2.send_keys(Keys.RETURN)
    time.sleep(2)
    try:
        browser.find_element_by_id('btnContinue').click()
    except:
        pass
    time.sleep(3)
    try:
        browser.find_element_by_xpath('/html/body/div/a').click()
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
    browser.get('https://ipon.sso.francetelecom.fr/NGI/GassiAccess.jsp')
    time.sleep(1)
    user_form = browser.find_element_by_id('user')
    user_form.send_keys('L')
    time.sleep(1)
    shell.SendKeys('{DOWN}', 0)
    time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    time.sleep(1)
    shell.SendKeys("{TAB}", 0)
    time.sleep(1)
    shell.SendKeys("{TAB}", 0)
    # elem2 = browser.find_element_by_id("password")
    # elem2.send_keys("Soge2017*")
    # time.sleep(1)
    shell.SendKeys("{ENTER}", 0)
    time.sleep(3)
    browser.get('https://ipon.sso.francetelecom.fr/NGI/GassiAccess.jsp')
    time.sleep(4)


def crear_proyecto_ipon(browser, nra, project):

    # Pulsar mon bureau
    browser.get('https://ipon.sso.francetelecom.fr/desktop.jsp')
    time.sleep(3)
    nouveau_project = browser.find_element_by_xpath('/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr/td[2]/a')
    nouveau_project.click()
    time.sleep(3)
    nom = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[1]/td[2]/input')
    nom.clear()
    nom.send_keys(project['nom_project'])
    code_secteur = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[5]/td[2]/input')
    code_secteur.clear()
    code_secteur.send_keys(nra)
    code_oeie = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[6]/td[2]/input')
    code_oeie.clear()
    code_oeie.send_keys('000000')
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[1]/tbody/tr/td[4]/a').click()
    time.sleep(2)


def select_imb(browser, imbs, inmueble):
    browser.find_element_by_xpath('/html/body/div/div[1]/table/tbody/tr[1]/td/table/tbody/tr/td[3]').click()
    time.sleep(1)
    research_immueble = browser.find_element_by_xpath('/html/body/div/div[1]/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/ul/li/ul/li[14]/a/span')
    research_immueble.click()
    time.sleep(1)
    id_inmuble_form = browser.find_element_by_xpath(
        '/html/body/div[1]/div[1]/div/form[1]/div/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[1]/table/tbody/tr/td[3]/font/div/input')
    time.sleep(1)
    id_inmuble_form.send_keys('I')
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


def estudio(browser, nra, imbs, inmueble):
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
    time.sleep(3)
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

    pass
def consulter_metre(browser, imbs, inmueble):
    # consulter metre la jour
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[1]/tbody/tr/td[4]/a').click()
    time.sleep(3)

    for batiment in imbs[inmueble].keys():
        # creer batiment
        # browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[10]/font/a[2]').click()
        # time.sleep(2)
        # browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/form/table/tbody/tr[3]/td/font/div[1]/a').click()
        time.sleep(3)
        # browser.set_window_position(WX, WY)
        # browser.set_window_size(1700, 1100)
        shell.SendKeys("{F12}", 0)
        time.sleep(6)
        win32_click(2397, 757)
        time.sleep(2)
        # type_batiment = browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]'
        #                                               '/td/form/table[2]/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr/td[6]/div')
        # elem_but_pos(type_batiment)
        #
        # time.sleep(1)
        # for i in range(4):
        #     shell.SendKeys('{UP}', 0)
        #     time.sleep(1)
        # shell.SendKeys('{DOWN}', 0)
        # time.sleep(1)
        # shell.SendKeys("{ENTER}", 0)
        # time.sleep(1)
        # etat_batiment = browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]'
        #                                               '/td/form/table[2]/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr/td[7]/div')
        # elem_but_pos(etat_batiment)
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
        time.sleep(5)
        # click [no name]
        browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/tabl'
                                      'e[2]/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr/td[2]/div/a').click()
        time.sleep(2)
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
        # #TODO: meter el batiment que corresponda
        niveles = imbs[inmueble][batiment].keys()

        lista_niveles = []
        for nivel in niveles:
            if nivel != 'bti':
                lista_niveles.append(int(nivel))
        sorted_niveles = sorted(lista_niveles)
        num_niveles = len(lista_niveles)
        nombre_form.send_keys(str(num_niveles))
        apartirde_form = browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/form/table/tbody/tr[2]/td[2]/font/input')
        apartirde_form.clear()
        apartirde_form.send_keys(str(sorted(lista_niveles)[0]))
        time.sleep(1)
        browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/form/table/tbody/tr[4]/td/font/div[1]/a').click()
        time.sleep(4)
        # type de lescalier
        browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[2]/font/a[2]').click()
        time.sleep(3)
        type_lescalier = browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/'
                                                       'tbody/tr[4]/td/form/table[2]/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr/td[4]/div')
        elem_but_pos(type_lescalier, False)
        time.sleep(1)
        shell.SendKeys('{DOWN}', 0)
        time.sleep(1)
        shell.SendKeys('{DOWN}', 0)
        time.sleep(1)
        shell.SendKeys("{ENTER}", 0)
        time.sleep(2)
        browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[2]/font/a[2]').click()
        time.sleep(3)

        browser.find_element_by_xpath(
            '/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[2]/font/a[2]').click()
        i = 1
        time.sleep(1)
        for nivel in sorted_niveles:
            time.sleep(1)
            if nivel == 0:
                nom_nivel = browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/'
                                                          'tr[1]/td/form/table[2]/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr[' + str(i) + ']/td[2]/div')
                elem_but_pos(nom_nivel)
                time.sleep(2)
                shell.Sendkeys("{LEFT}", 0)
                time.sleep(1)
                shell.Sendkeys("{DELETE}", 0)
                time.sleep(1)
                shell.SendKeys("RCD", 0)
                time.sleep(1)
                shell.SendKeys("{ENTER}", 0)
            i += 1

        # Conexiones de cada piso
        browser.find_element_by_xpath(
            '/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[2]/font/a[2]').click()
        time.sleep(4)
        i = 1
        for nivel in sorted_niveles: #Numero de niveles
            time.sleep(1)
            num_conexiones = browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr[' + str(i) + ']/td[4]/div')
            elem_but_pos(num_conexiones)
            time.sleep(2)
            shell.SendKeys(imbs[inmueble][batiment][str(nivel)]['num_el'], 0)
            time.sleep(1)
            shell.SendKeys("{ENTER}", 0)
            time.sleep(2)
            i += 1
        i = 1
        browser.find_element_by_xpath(
            '/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[2]/font/a[2]').click()
        time.sleep(4)
        for nivel in sorted_niveles: #Numero de niveles

            time.sleep(1)
            type_level = browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr[' + str(i) + ']/td[6]/div')
            elem_but_pos(type_level,False)
            time.sleep(2)
            if imbs[inmueble][batiment][str(nivel)]['num_el'] != '0' or imbs[inmueble][batiment][str(nivel)]['material']:
                shell.SendKeys('{DOWN}', 0)
                time.sleep(1)
                if imbs[inmueble][batiment][str(nivel)]['material']:
                    shell.SendKeys('{DOWN}', 0)
                    time.sleep(1)
                    shell.SendKeys('{DOWN}', 0)
                    time.sleep(1)
                    if imbs[inmueble][batiment][str(nivel)]['num_el'] == '0':
                        shell.SendKeys('{DOWN}', 0)
                        time.sleep(1)
            shell.SendKeys("{ENTER}", 0)
            time.sleep(2)
            i += 1
        time.sleep(2)
        # guardar
        browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[2]/font/a[2]').click()
        time.sleep(4)
        browser.find_element_by_xpath('/html/body/table[2]/tbody/tr/td[2]/table/tbody/tr[1]/td/div/a[6]').click()
        time.sleep(3)
        browser.find_element_by_xpath('/html/body/table[2]/tbody/tr/td[2]/table/tbody/tr[2]/td/p/nobr[3]/a').click()
        time.sleep(3)
        # crear acces
        browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[5]/font/a[2]').click()
        time.sleep(2)
        #crear contact
        browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[5]/font/a[2]').click()
        time.sleep(2)
        browser.find_element_by_xpath('/html/body/table[2]/tbody/tr/td[2]/table/tbody/tr[1]/td/div/a[6]').click()
        time.sleep(3)
        browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr/td[2]/div/a').click()
        time.sleep(3)
        browser.find_element_by_xpath('/html/body/table[2]/tbody/tr/td[2]/table/tbody/tr[2]/td/p/nobr[2]/a').click()
        time.sleep(3)
        type_access = browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/'
                                                    'table[2]/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr/td[9]/div')
        elem_but_pos(type_access)
        time.sleep(3)
        shell.SendKeys('{DOWN}', 0)
        time.sleep(1)
        shell.SendKeys('{DOWN}', 0)
        time.sleep(1)
        shell.SendKeys("{ENTER}", 0)
        time.sleep(2)
        browser.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/form/table[2]/tbody/tr/td/table[1]/tbody/tr/td/table/tbody/tr/td[2]/font/a[2]').click()
    time.sleep(3)


def crear_pb(browser, imbs, inmueble, pbs, btis):
    pythoncom.CoInitialize()
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(r'C:\Users\josko\PycharmProjects\josko\cablage4.xlsx')
    excel.Visible = True
    sheet = wb.Worksheets(1)
    canvas = sheet.Shapes
    time.sleep(8)
    for pb in pbs:
        if pbs[pb]['inmueble'] == inmueble:
            browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form[2]/table/tbody/tr/td/table/tbody/tr/td[5]/a').click()
            time.sleep(2)
            rechercher_button = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[5]/td/div[2]/div/a')
            rechercher_button.click()
            time.sleep(1)
            select_modele = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/table/tbody/tr/'
                                                          'td/table/tbody/tr[6]/td[2]/select').click()
            if pbs[pb]['tipo'] == '[6.i.aco]':
                selection_modele = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/'
                                                                 'table/tbody/tr[6]/td[2]/select/option[12]').click()
            if pbs[pb]['tipo'] == '[12.i13.3m]':
                selection_modele = browser.find_element_by_xpath('/html/body/div/div[1]/div/form/table/tbody/tr/td/'
                                                                     'table/tbody/tr[6]/td[2]/select/option[5]').click()
            time.sleep(1)
            # click creer
            browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[17]/td/div[1]/div/a').click()
            #TODO hacer lo de añadir mas cassetes si es necesario
            time.sleep(3)
            browser.find_element_by_xpath('/html/body/div[1]/div[1]/table/tbody/tr[3]/td/table/tbody/tr/td/div[4]/table/tbody/tr/td[2]/a').click()
            time.sleep(2)
            browser.find_element_by_xpath('/html/body/div/div[1]/div/div[3]/form/table/tbody/tr/td/table[1]/tbody/tr/td[2]/a').click()
            time.sleep(2)
            pt_value = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table'
                                                     '[2]/tbody/tr[1]/td[2]/input').get_attribute('value')
            for shp in canvas:
                box = shp.TextFrame2.TextRange.Characters.Text
                if 'PT ' + pb in box:
                    shp.TextFrame2.TextRange.Characters.Text = pt_value

            pbs[pb]['pt'] = pt_value
            #Hauteur par rapport au sol
            browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[15]/td[2]/select').click()
            browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[15]/td[2]/select/option[2]').click()
            #Position lequipament
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[16]/td[2]/select').click()
            if '3M (GT)' in pbs[pb]['observation']:
                browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]'
                                              '/tbody/tr[16]/td[2]/select/option[10]').click()
            else:
                browser.find_element_by_xpath(
                    '/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[16]/td[2]/select'
                    '/option[6]').click()
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[1]/tbody/tr/td[2]/a').click()
            # Volver al inmueble
            time.sleep(2)
            browser.find_element_by_xpath('/html/body/div/div[1]/table/tbody/tr[3]/td/div/div/a[7]').click()
            time.sleep(2)

    for bti in btis:
        if btis[bti]['imb_is'] == inmueble:
            browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/div/form[2]/table/tbody/tr/td/table/tbody/tr/td[5]/a').click()
            time.sleep(2)
            rechercher_button = browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[5]/td/div[2]/div/a')
            rechercher_button.click()
            time.sleep(1)
            select_modele = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/table/tbody/tr/'
                                                          'td/table/tbody/tr[6]/td[2]/select').click()
            if btis[bti]['tipo'] == 'BTI 36':
                selection_modele = browser.find_element_by_xpath(
                    '/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[6]/td[2]/select/option[19]').click()
            if btis[bti]['tipo'] == 'BTI 144':
                selection_modele = browser.find_element_by_xpath(
                    '/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[6]/td[2]/select/option[21]').click()
            time.sleep(1)
            # click creer
            browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[17]/td/div[1]/div/a').click()
            # TODO hacer lo de añadir mas cassetes si es necesario
            time.sleep(3)
            browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/table/tbody/tr[3]/td/table/tbody/tr/td/div[4]/table/tbody/tr/td[2]/a').click()
            time.sleep(2)
            browser.find_element_by_xpath(
                '/html/body/div/div[1]/div/div[3]/form/table/tbody/tr/td/table[1]/tbody/tr/td[2]/a').click()
            time.sleep(2)
            pt_value = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table'
                                                     '[2]/tbody/tr[1]/td[2]/input').get_attribute('value')
            pt_split = pt_value.split(sep=' ')
            pt_value = pt_split[1]
            for shp in canvas:
                box = shp.TextFrame2.TextRange.Characters.Text
                if 'bti' + bti in box and 'BTI ' in box:
                    box = box.replace('bti'+bti, pt_value)
                    shp.TextFrame2.TextRange.Characters.Text = box

            btis[bti]['pt'] = pt_value
            # Hauteur par rapport au sol
            browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[15]/td[2]/select').click()
            browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[15]/td[2]/select/option[2]').click()
            # Position lequipament
            time.sleep(1)
            browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[16]/td[2]/select').click()
            if '36 (GT)' in pbs[pb]['observation'] or '144 (GT)' in pbs[pb]['observation']:
                browser.find_element_by_xpath(
                    '/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]'
                    '/tbody/tr[16]/td[2]/select/option[10]').click()
            else:
                browser.find_element_by_xpath(
                    '/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[2]/tbody/tr[16]/td[2]/select'
                    '/option[6]').click()
            time.sleep(1)
            browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/div/div[3]/form/table/tbody/tr/td/table[1]/tbody/tr/td[2]/a').click()
            # Volver al inmueble
            time.sleep(2)
            browser.find_element_by_xpath('/html/body/div/div[1]/table/tbody/tr[3]/td/div/div/a[7]').click()
            time.sleep(2)

    for pb in pbs:
        if pbs[pb]['inmueble'] == inmueble:
            #CH 01573 de ejemplo luego se pillará de un excel
            browser.find_element_by_xpath('/html/body/div/div[1]/table/tbody/tr[1]/td/table/tbody/tr/td[3]').click()
            time.sleep(1)
            rechercher_chambre = browser.find_element_by_xpath('/html/body/div[1]/div[1]/table/tbody/tr[1]/td/table/'
                                                               'tbody/tr/td[3]/div/ul/li/ul/li[16]/a/span')
            rechercher_chambre.click()
            time.sleep(2)
            input_chambre = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form[1]/div/table/tbody/tr[2]/'
                                                          'td[1]/table/tbody/tr[3]/td/div/table/tbody/tr/td[3]/font/span/input')
            chambre_code = '01573'
            pt_code = 'PT 1394'
            time.sleep(1)
            input_chambre.send_keys(chambre_code)
            shell.SendKeys("{ENTER}", 0)
            time.sleep(4)
            insee = inmueble.split('/')[1]
            identifiant = chambre_code + '/' + insee
            row_ch = browser.find_element_by_xpath("//*[contains(text(), '" + identifiant + "')]")
            row_parent = row_ch.find_element_by_xpath('..')
            ch = row_parent.find_element_by_xpath('td[2]').click()
            time.sleep(2)
            pt_selected = browser.find_element_by_xpath("//*[contains(text(), '" + pt_code + "')]")
            row_parent2 = pt_selected.find_element_by_xpath('../..')
            pa_selected = row_parent2.find_element_by_xpath('td[4]/a').click()
            time.sleep(2)
            creer_point = browser.find_element_by_xpath('/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/'
                                                        'tr/td[2]/a').click()
            time.sleep(1)
            main_window = browser.current_window_handle
            mas1 = browser.find_element_by_xpath('/html/body/div/div[1]/div/form/table[1]/tbody/tr/td/table/thead/t'
                                                 'r[4]/td[2]/b[1]/a').click()
            time.sleep(4)
            signin_window_handle = browser.window_handles[1]
            time.sleep(1)
            browser.switch_to.window(signin_window_handle)
            frame1 = browser.find_element_by_xpath('/html/frameset/frame[1]')
            frame2 = browser.find_element_by_xpath('/html/frameset/frame[2]')
            browser.switch_to_frame(frame1)
            # Seleccionar inmueble
            browser.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/ul/li/div/form/input[1]').click()
            time.sleep(1)
            research_inmueble = browser.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/ul/li/ul/li[14]/a')
            research_inmueble.click()
            time.sleep(1)
            id_inmuble_form = browser.find_element_by_xpath(
                '/html/body/div[1]/div/div/form[1]/div/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[1]/table/tbody/tr/td[3]/font/div/input')
            time.sleep(1)
            id_inmuble_form.send_keys('I')
            shell.SendKeys('{DOWN}', 0)
            time.sleep(1)
            shell.SendKeys("{ENTER}", 0)
            time.sleep(1)
            shell.SendKeys("{ENTER}", 0)
            time.sleep(8)
            browser.find_element_by_xpath(
                '/html/body/div[1]/div/div/form[2]/table/tbody/tr/td/div[1]/table/tbody/tr/td[1]/input').click()
            time.sleep(1)
            browser.switch_to_default_content()
            time.sleep(1)
            browser.switch_to_frame(frame2)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/form/div[1]/div/a').click()
            time.sleep(2)
            browser.switch_to_window(main_window)
            time.sleep(1)
            main_window = browser.current_window_handle
            mas2 = browser.find_element_by_xpath('/html/body/div/div[1]/div/form/table[1]/tbody/tr/td/table/thead/tr'
                                                 '[5]/td[2]/b[1]/a').click()
            time.sleep(4)
            signin_window_handle = browser.window_handles[1]
            browser.switch_to.window(signin_window_handle)
            time.sleep(1)
            frame1 = browser.find_element_by_xpath('/html/frameset/frame[1]')
            frame2 = browser.find_element_by_xpath('/html/frameset/frame[2]')
            browser.switch_to_frame(frame1)
            i = 2
            for k in range(1):
                browser.find_element_by_xpath('/html/body/div/div/div/form/table/tbody/tr[2]/td/table/tbody/tr[' + str(i) + ']/td[1]/input').click()
                time.sleep(1)
                i += 1
            time.sleep(1)
            browser.switch_to_default_content()
            time.sleep(1)
            browser.switch_to_frame(frame2)
            browser.find_element_by_xpath('/html/body/form/div[1]/div/a').click()
            time.sleep(2)
            browser.switch_to_window(main_window)
            browser.find_element_by_xpath('/html/body/div/div[1]/div/form/table[2]/tbody/tr/td/a[1]').click()
    wb.Close(True)


def crear_cables(browser, imbs, inmueble, pbs, btis, cables, pa_chambre):
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/table/tbody/tr[3]/td/table/tbody/tr/td/div[2]/table/tbody/tr/td[2]/a').click()
    time.sleep(2)

    for cable in cables:
        if len(cables[cable]['ini']) == 1 or len(cables[cable]['ini']) > 2:  # si el cable va de pb a pb o de bti a pb
            time.sleep(3)
            browser.find_element_by_xpath(
                '/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/table/tbody/tr/td[4]/a').click()
            time.sleep(1)
            num_fo_form = browser.find_element_by_xpath(
                '/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[3]/td[2]/input')
            num_fo_form.clear()
            num_fo_form.send_keys(cables[cable]['num_fo'])
            time.sleep(1)
            # Pulsar el + para añadir site suport que conecta los cables
            browser.find_element_by_xpath(
                '/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[8]/td[2]/a[1]').click()
            time.sleep(4)
            main_window = browser.current_window_handle
            time.sleep(2)
            signin_window_handle = browser.window_handles[1]
            browser.switch_to.window(signin_window_handle)
            time.sleep(3)
            frame = browser.find_element_by_xpath('/html/frameset/frame[1]')
            frame2 = browser.find_element_by_xpath('/html/frameset/frame[2]')
            browser.switch_to_frame(frame)
            time.sleep(2)
            #TODO quitar los pts de los pbs y btis por defecto
            for pb in pbs:
                if cables[cable]['ini'] == pb:
                    if pbs[pb]['inmueble'] == inmueble:
                        pbs[pb]['pt'] = '002586'
                        select_pt_in_imb(browser, frame2, pbs[pb]['pt'])
            for bti in btis:
                ini = cables[cable]['ini']
                if ini[-1] == bti:
                    if btis[bti]['inmueble'] == inmueble:
                        btis[bti]['pt'] = '002586'
                        select_pt_in_imb(browser, frame2, btis[bti]['pt'])

            browser.switch_to_window(main_window)
            time.sleep(1)
            browser.find_element_by_xpath(
                '/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[8]/td[2]/a[1]').click()
            time.sleep(4)
            main_window = browser.current_window_handle
            time.sleep(2)
            signin_window_handle = browser.window_handles[1]
            browser.switch_to.window(signin_window_handle)
            time.sleep(3)
            frame = browser.find_element_by_xpath('/html/frameset/frame[1]')
            frame2 = browser.find_element_by_xpath('/html/frameset/frame[2]')
            browser.switch_to_frame(frame)
            time.sleep(2)
            # TODO quitar los pts de los pbs y btis por defecto
            for pb in pbs:
                if cables[cable]['fin'] == pb:
                    if pbs[pb]['inmueble'] == inmueble:
                        pbs[pb]['pt'] = '002586'
                        select_pt_in_imb(browser, frame2, pbs[pb]['pt'])

            browser.switch_to_window(main_window)
            time.sleep(2)
            # Crear cable
            browser.find_element_by_xpath(
                '/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[15]/td/a[1]').click()

    for cable in cables:
        # Click en la comuna ej. "picardie"
        browser.find_element_by_xpath('/html/body/div[1]/div[1]/table/tbody/tr[3]/td/div/div/a[4]').click()
        time.sleep(2)
        browser.find_element_by_xpath('/html/body/div[1]/div[1]/table/tbody/tr[3]/td/table/tbody/tr/td/div[2]/table/tbody/tr/td[2]/a').click()
        # if cables[cable]['ini'] == 'PA':  # si el cable es externo
            # time.sleep(1)
            # browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/table/tbody/tr/td[4]/a').click()
            # time.sleep(1)
            # num_fo_form = browser.find_element_by_xpath('/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[3]/td[2]/input')
            # num_fo_form.clear()
            # num_fo_form.send_keys(cables[cable]['num_fo'])
            # time.sleep(1)
            # # Pulsar el + para añadir site suport que conecta los cables
            # browser.find_element_by_xpath('/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[8]/td[2]/a[1]').click()
            # time.sleep(4)
            # main_window = browser.current_window_handle
            # time.sleep(2)
            # signin_window_handle = browser.window_handles[1]
            # browser.switch_to.window(signin_window_handle)
            # time.sleep(3)
            # frame = browser.find_element_by_xpath('/html/frameset/frame[1]')
            # frame2 = browser.find_element_by_xpath('/html/frameset/frame[2]')
            # browser.switch_to_frame(frame)
            # time.sleep(2)
            # select_pa(browser, pa_chambre, inmueble, frame2)
            # browser.switch_to_window(main_window)
            # time.sleep(1)
            # browser.find_element_by_xpath(
            #     '/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[8]/td[2]/a[1]').click()
            # time.sleep(4)
            # main_window = browser.current_window_handle
            # time.sleep(2)
            # signin_window_handle = browser.window_handles[1]
            # browser.switch_to.window(signin_window_handle)
            # time.sleep(3)
            # frame = browser.find_element_by_xpath('/html/frameset/frame[1]')
            # frame2 = browser.find_element_by_xpath('/html/frameset/frame[2]')
            # browser.switch_to_frame(frame)
            # time.sleep(2)
            # for pb in pbs:
            #     if cables[cable]['fin'] == pb:
            #         if pbs[pb]['inmueble'] == inmueble:
            #             select_imb_con_pt(browser, inmueble, frame2)
            # for bti in btis:
            #     fin = cables[cable]['fin']
            #     if fin[-1] == bti:
            #         if btis[bti]['imb_is'] == inmueble:
            #             select_imb_con_pt(browser, inmueble, frame2)
            # browser.switch_to_window(main_window)
            # time.sleep(2)
            # # Crear cable
            # browser.find_element_by_xpath('/html/body/div/div[1]/div/form/table/tbody/tr/td/table/tbody/tr[15]/td/a[1]').click()

        time.sleep(2)
        # Pulsar el cable
        browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/div[1]/table/tbody/tr/td[2]/a').click()
        time.sleep(3)
        # browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/div[1]/table/tbody/tr[1]/td[1]/input').click()
        # time.sleep(1)
        # browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/div[1]/table/thead/tr/th[3]/table/tbody/tr/td[3]/a').click()
        # time.sleep(4)
        # main_window = browser.current_window_handle
        # time.sleep(2)
        # signin_window_handle = browser.window_handles[1]
        # browser.switch_to.window(signin_window_handle)
        # time.sleep(3)
        # frame = browser.find_element_by_xpath('/html/frameset/frame[1]')
        # frame2 = browser.find_element_by_xpath('/html/frameset/frame[2]')
        # browser.switch_to_frame(frame)
        # time.sleep(2)
        # browser.find_element_by_xpath('/html/body/div[1]/div/div/form/table/tbody/tr/td/div[1]/table/tbody/tr/td[2]/a').click()
        # time.sleep(3)
        # browser.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td/div[2]/table/tbody/tr/td[2]/a').click()
        # time.sleep(2)
        # i = 1
        # while True:
        #     b = browser.find_element_by_xpath(
        #         '/html/body/div/div/div/form/table/tbody/tr/td/div[1]/table/tbody/tr[' + str(i) + ']/td[3]').text
        #     if b == '' and browser.find_element_by_xpath(
        #                             '/html/body/div/div/div/form/table/tbody/tr/td/div[1]/table/tbody/tr[' + str(
        #                             i) + ']/td[4]').text == 'Sortie':
        #         browser.find_element_by_xpath(
        #             '/html/body/div/div/div/form/table/tbody/tr/td/div[1]/table/tbody/tr[' + str(i) + ']/td[1]').click()
        #         break
        #     i += 1
        # time.sleep(2)
        # browser.switch_to_default_content()
        # time.sleep(1)
        # browser.switch_to_frame(frame2)
        # time.sleep(1)
        # browser.find_element_by_xpath('/html/body/form/div[1]/div/a').click()
        # time.sleep(1)
        # browser.switch_to_window(main_window)
        # time.sleep(3)
        # browser.find_element_by_xpath(
        #     '/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/div[1]/table/tbody/tr[3]/td[1]/input').click()
        # time.sleep(1)
        # browser.find_element_by_xpath(
        #     '/html/body/div[1]/div[1]/div/form/table/tbody/tr/td/div[1]/table/thead/tr/th[3]/table/tbody/tr/td[3]/a').click()
        # time.sleep(4)
        # main_window = browser.current_window_handle
        # time.sleep(2)
        # signin_window_handle = browser.window_handles[1]
        # browser.switch_to.window(signin_window_handle)
        # time.sleep(3)
        # frame = browser.find_element_by_xpath('/html/frameset/frame[1]')
        # frame2 = browser.find_element_by_xpath('/html/frameset/frame[2]')
        # browser.switch_to_frame(frame)
        # time.sleep(2)
        # browser.find_element_by_xpath(
        #     '/html/body/div[1]/div/div/form/table/tbody/tr/td/div[1]/table/tbody/tr/td[2]/a').click()
        # time.sleep(3)
        # browser.find_element_by_xpath(
        #     '/html/body/div[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td/div[2]/table/tbody/tr/td[2]/a').click()
        # time.sleep(2)
        # i = 1
        # while True:
        #     b = browser.find_element_by_xpath(
        #         '/html/body/div/div/div/form/table/tbody/tr/td/div[1]/table/tbody/tr[' + str(i) + ']/td[3]').text
        #     if b == '' and browser.find_element_by_xpath(
        #                             '/html/body/div/div/div/form/table/tbody/tr/td/div[1]/table/tbody/tr[' + str(
        #                         i) + ']/td[4]').text == 'Entrée':
        #         browser.find_element_by_xpath(
        #             '/html/body/div/div/div/form/table/tbody/tr/td/div[1]/table/tbody/tr[' + str(i) + ']/td[1]').click()
        #         break
        #     i += 1
        # time.sleep(2)
        # browser.switch_to_default_content()
        # time.sleep(1)
        # browser.switch_to_frame(frame2)
        # time.sleep(1)
        # browser.find_element_by_xpath('/html/body/form/div[1]/div/a').click()
        # time.sleep(1)
        # browser.switch_to_window(main_window)
        # time.sleep(3)
        tr_number = browser.find_element_by_xpath('/html/body/div[1]/div[1]/table/tbody/tr[3]/td/div/div/a[6]').text
        # pythoncom.CoInitialize()
        # excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
        # wb = excel.Workbooks.Open(r'C:\Users\josko\PycharmProjects\josko\cablage4.xlsx')
        # excel.Visible = True
        # sheet = wb.Worksheets(1)
        # canvas = sheet.Shapes
        # time.sleep(8)
        cables[cable]['nombre'] = tr_number
        # for shp in canvas:
        #     box = shp.TextFrame2.TextRange.Characters.Text
        #     if 'TR' in box:
        #         box = box.replace('TR PA-a', tr_number)
        #         shp.TextFrame2.TextRange.Characters.Text = box
        # time.sleep(1)
        # Volver a cables
        browser.find_element_by_xpath('/html/body/div[1]/div[1]/table/tbody/tr[3]/td/div/div/a[5]').click()
        time.sleep(2)
        # browser.set_window_position(WX, WY)
        # browser.set_window_size(1700, 1100)
        time.sleep(1)
    for cable in cables:
        cable_nom = cables[cable]['nombre']
        cn1 = cable_nom[3:5]
        cn2 = cable_nom[-4:]
        row_cable = browser.find_element_by_xpath("//*[contains(text(), '" + cn1 + "') and contains(text(), '" + cn2 + "')]")
        row_parent = row_cable.find_element_by_xpath('../..')
        long_cable = row_parent.find_element_by_xpath('td[3]')
        elem_but_pos(long_cable)


def ejecutar_ipon(nra):

    project = get_data()
    imbs = project ['inmuebles']
    pbs = project ['pbs']
    btis = project ['btis']
    cables = project['cables']
    pa_chambre = project['pa_chambre']
    pa_pt = project['pa_pt']
    browser = set_up_browser()
    login(browser)
    # crear_proyecto_ipon(browser, nra, project)
    for inmueble in imbs.keys():
        select_imb(browser, imbs, inmueble)
        # estudio(browser, nra, imbs, inmueble)
        # consulter_metre(browser, imbs, inmueble)
        # crear_pb(browser, imbs, inmueble, pbs, btis)
        crear_cables(browser, imbs, inmueble, pbs, btis, cables, pa_chambre)


# ejecutar_ipon()
