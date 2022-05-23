from selenium import webdriver
import os
import pickle
import Remedy_Web
import Remedy_Web_1_0
import time
import sys
import win32com.client
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from retry import retry
import xlrd,win32com.client
import datetime
from docxtpl import DocxTemplate

def login(driver,login,passw):
    driver.get("https://cedar.mts.ru/")
    try:
        login_el = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "loginform-login")))
        # login_el=driver.find_element_by_id('loginform-login')
        login_el.send_keys(login)
        passw_el=driver.find_element_by_id('loginform-password')
        passw_el.send_keys(passw)
        ok_button=driver.find_element_by_name('login-button')
        ok_button.click()
    except:
        try:
            login_el = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="main-page-container"]/div/div[2]/form/div/div[1]/div/inputn')))
            return 'ok'
        except:
            login(driver, login, passw)



def openBS(driver, BSID):
    driver.get("https://cedar.mts.ru/mrnew/bs/?bsId="+BSID)

def isMBHexist(driver):
    nioss_el = driver.find_elements_by_xpath('/html/body/table/tbody/tr/td[2]/div[7]/table/tbody/tr[5]/td/table/tbody/tr/td/a/b')
    for el in nioss_el:
        if "MBH" in el.get_attribute('innerText'):
            return True
    return False

def PL_to_Cedar(PL):
    try:
        BS = PL.split('_')[1] + '-' + PL.split('_')[2].rjust(5, '0')
        return BS
    except:
        return False

def set_value(driver,Name,value):
    try:
        if value:
            el = driver.find_element_by_name(Name)
            el.clear()
            el.send_keys(value)
        return True
    except:
        return False

def sel_value(driver,Name,value):
    try:
        if value:
            el = driver.find_element_by_name(Name)
            # el.click()
            el.send_keys(value.strip())
            el.send_keys(Keys.ENTER)
        return True
    except:
        return False


def addRRL(driver,RRL_Row):
    BS_A = PL_to_Cedar(RRL_Row[0])
    BS_Z = PL_to_Cedar(RRL_Row[1])
    if BS_A and BS_Z:
        driver.get('https://cedar.mts.ru/mrnew/prolet/?act=newProlet&prlType=rrl&aNumber=%s' %BS_A)
        set_value(driver,"bNumber",BS_Z)
        el=driver.find_elements_by_xpath('//*[@id="but"]')
        el[3].click()
        time.sleep(4)
        el = driver.find_elements_by_xpath('//*[@id="but"]')
        el[3].click()
        time.sleep(4)
        # создали пролет, теперь его заполняем
        sel_value(driver, "techReshenie", RRL_Row[2])#тех. решение
        sel_value(driver, "aTrans", RRL_Row[3])# Оборудование А тип корзины
        sel_value(driver, "bTrans", RRL_Row[4])# Оборудование Б тип корзины
        sel_value(driver, "range", RRL_Row[5])# Диапазон
        sel_value(driver, "aArrlType1", RRL_Row[6])# Диаметр А
        sel_value(driver, "bArrlType1", RRL_Row[7])# Диаметр Б
        sel_value(driver, "rezerv", RRL_Row[8])# Конфигурация
        set_value(driver, "niossName", RRL_Row[9])# Название в НИОСС
        set_value(driver, "power", RRL_Row[10])# Мощность на выходе
        set_value(driver, "datePlan", RRL_Row[11])# Плановая дата
        set_value(driver, "notes", RRL_Row[12])# Примечание
        sel_value(driver, "source", RRL_Row[13])# источник оборудования
        sel_value(driver, "polarization", RRL_Row[14])# Поляризация
        sel_value(driver, "bandWidth", RRL_Row[15])# Ширина полосы
        time.sleep(1)
        set_value(driver, "curator", RRL_Row[16])# Куратор
        time.sleep(1)
        set_value(driver, "hyperion", '0010600210630001')  # Проект
        if "XPIC" in RRL_Row[8]:
            sel_value(driver, "combaner", 'Симметричный')

        el=driver.find_elements_by_xpath('//*[@id="but"]')
        # for e in el[3:-1]:
        el[3].click()
    time.sleep(8)
    print(driver.current_url)

def addMBH(driver,BSID,MBH_name,MBH_type,date,IP,notes):
    driver.get('https://cedar.mts.ru/mrnew/prolet?act=addRack&bsId=%s' %BSID)
    el=driver.find_element_by_name("rackName")
    el.send_keys(MBH_name)
    el=driver.find_element_by_xpath('/html/body/table/tbody/tr/td[2]/div/table/tbody/tr[4]/td[1]/table/tbody/tr[2]/td/span/span[1]/span').click()
    el = driver.find_element_by_xpath('/html/body/span/span/span[1]/input')
    el.send_keys(str(MBH_type))
    el.send_keys(Keys.ENTER)
    el = driver.find_element_by_name("rackLevel")
    el.send_keys("узел MPLS")

    el = driver.find_element_by_name("mux")
    el.send_keys("10G")

    el = driver.find_element_by_name("hyperion")
    el.send_keys(IP)

    el = driver.find_element_by_name("niossName")
    el.send_keys(MBH_name)

    el = driver.find_element_by_name("notes")
    el.send_keys(notes)
    el = driver.find_element_by_name("orgId")
    el.send_keys("Связь-монтаж ЗАО")


    el = driver.find_element_by_name("datePlan")
    el.clear()
    el.send_keys(date)

    el=driver.find_elements_by_xpath('//*[@id="but"]')
    # for e in el[3:-1]:
    el[3].click()
    time.sleep(6)
    print("-->"+driver.current_url)

def editMBH(driver,MBHID,MBH_name):
    driver.get('https://cedar.mts.ru/mrnew/prolet/?act=loadRack&rackId=%s' %MBHID)
    el=driver.find_element_by_name("niossName")
    el.send_keys(MBH_name)
    # el=driver.find_element_by_xpath('/html/body/table/tbody/tr/td[2]/div/table/tbody/tr[5]/td[1]/table/tbody/tr/td[2]/table/tbody/tr[5]/td/q').click()
    # el = driver.find_element_by_xpath('/html/body/span/span/span[1]/input')
    # el.send_keys(MBH_type)
    # el.send_keys(Keys.ENTER)
    # el = driver.find_element_by_name("rackLevel")
    # el.send_keys("узел MPLS")
    #
    # el = driver.find_element_by_name("mux")
    # el.send_keys("10G")
    #
    # el = driver.find_element_by_name("hyperion")
    # el.send_keys("0010600200610002")
    #
    # el = driver.find_element_by_name("niossName")
    # el.send_keys(MBH_name)
    #
    # el = driver.find_element_by_name("notes")
    # el.send_keys("Phase3 2020г. ЮГ 01/23665с от 26.12.2019")
    # el = driver.find_element_by_name("orgId")
    # el.send_keys("Интраком Связь ООО")
    #
    #
    # el = driver.find_element_by_name("datePlan")
    # el.clear()
    # el.send_keys(date)

    el=driver.find_elements_by_xpath('//*[@id="but"]')
    # for e in el[3:-1]:
    el[3].click()
    time.sleep(6)
    print(driver.current_url)

def addMBHlist():
    rb = xlrd.open_workbook('C:/Users/yvsobol3/PycharmProjects/Test/MBH Cedar.xlsx')
    sheet = rb.sheet_by_index(0)
    start_time = time.time()

    options = webdriver.ChromeOptions()
    # options.add_argument('headless')  # если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    driver = webdriver.Chrome(options=options)
    login(driver, Remedy_Web.Credentials['login'], Remedy_Web.Credentials['pass'])
    time.sleep(6)
    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        if row[0] and row[1]:
            openBS(driver, str(row[4]))
            # if isMBHexist(driver):
            #     print("Узел на площадке %s существует" % row[0])
            # else:
            addMBH(driver, row[4], row[3], row[2], row[1])

def addRRLlist():
    rb = xlrd.open_workbook('C:/Users/yvsobol3/PycharmProjects/Test/MW Cedar all.xlsx')
    sheet = rb.sheet_by_index(0)
    start_time = time.time()

    options = webdriver.ChromeOptions()
    # options.add_argument('headless')  # если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    driver = webdriver.Chrome(options=options)
    login(driver, Remedy_Web.Credentials['login'], Remedy_Web.Credentials['pass'])
    time.sleep(5)
    for rownum in range(sheet.nrows)[1:]:
        row = sheet.row_values(rownum)
        if row[0] and row[1]:
            # openBS(driver,getCedarBSID(row[0],''))
            # if isMBHexist(driver):
            #     print("Узел на площадке %s существует" % row[0])
            # else:
            addRRL(driver, row[0:18])


def editMBHlist():
    rb = xlrd.open_workbook('C:/Users/yvsobol3/PycharmProjects/Test/MBH Cedar NIOSSID.xlsx')
    sheet = rb.sheet_by_index(0)
    start_time = time.time()

    options = webdriver.ChromeOptions()
    options.add_argument('headless')  # если хотим запустить chrome недивимкой
    options.add_argument("window-size=1920x1080")
    driver = webdriver.Chrome(options=options)
    login(driver, Remedy_Web.Credentials['login'], Remedy_Web.Credentials['pass'])
    time.sleep(6)
    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        if row[0] and row[1]:
            editMBH(driver, row[0], row[1])

def getCedarBSID(driver,PL,to_name):

    def getBSIDfromURL(url):
        try:
            BSID=url.split('bsId=')[1]
        except:
            BSID =''
        return BSID

    PL=PL.replace('_','')
    time.sleep(1)
    driver.get('https://cedar.mts.ru/basic/web/bs/bs/search-by-number?nType=context&bsNumber='+ PL +'&x=0&y=0')
    # login_el = WebDriverWait(driver, 10).until(
    #     EC.url_to_be((By.XPATH, '//*[@id="main-page-container"]/div/div[2]/form/div/div[1]/div/inputn')))
    time.sleep(1)
    url=driver.current_url
    if 'bsId=' in url:
        return getBSIDfromURL(url)
    else:
        els_type=driver.find_elements_by_xpath('//*[@id="crud-datatable-container"]/table/tbody/tr/td[2]')
        if els_type:
            for n,el_type in enumerate(els_type):
                if 'МБ' in el_type.text:
                    url=driver.find_element_by_xpath('//*[@id="crud-datatable-container"]/table/tbody/tr['+str(n+1)+']/td[4]/a').get_attribute('href')
                    return getBSIDfromURL(url)
        else:
            reply_send(to_name, '', 'Автоматические исходные данные', 'БС ' + PL + ' не найдена в Кедр', "")
            sys.exit()

    # rb = xlrd.open_workbook('C:/Users/yvsobol3/PycharmProjects/Test/bs_number_to_nioss_name_.xlsx')
    # sheet = rb.sheet_by_index(0)
    # start_time = time.time()
    # # time.sleep(6)
    # for rownum in range(9, sheet.nrows):
    #     row = sheet.row_values(rownum)
    #     if row[0] and row[1]:
    #         if "PL_"+PL == row[1]:
    #             # print("Время выполнения поиска БС по файлу: %s секунд" % (time.time() - start_time))
    #             return row[3]
    # # print("Время выполнения поиска БС по файлу: %s секунд" % (time.time() - start_time))
    # reply_send(to_name, '', 'Автоматические исходные данные', 'БС '+PL +' не найдена в Кедр', "")
    # sys.exit()

def reply_send(to_address,to_name,subject,text,ATT_PATH):
    # инициализируем объект outlook
    olk = win32com.client.Dispatch("Outlook.Application")
    Msg = olk.CreateItem(0)

    # формируем письма, выставляя адресата, тему и текст
    Msg.To = to_address
    Msg.Subject =  subject  # добавляем RE в тему
    if ATT_PATH:
        for PATH in ATT_PATH:
            Msg.Attachments.Add(PATH)
    Msg.Body = text
        # и отправляем
    Msg.Send()

def get_data_PL(driver,PL,to_name):
    BSID = getCedarBSID(driver,PL,to_name)
    driver.get('https://cedar.mts.ru/mrnew/bs/?bsId=%s' % BSID)

    # Адрес А
    Address = driver.find_element_by_xpath(
        '/html/body/table/tbody/tr/td[2]/div[1]/div[2]/table/tbody/tr[1]/td/div/div/div/div/div[1]/div/div/div/div[2]/div[1]/span').get_attribute(
        'innerText')
    try:
        AMS = driver.find_element_by_xpath(
            '//*[@id="mrs-bs-ams"]/div/div[2]/div/div/div/div[2]/div/span').get_attribute(
            'innerText')
    except:
        AMS = "Нет данных по опоре"
    try:
        Coord = driver.find_element_by_xpath(
            '/html/body/table/tbody/tr/td[2]/div[1]/div[2]/table/tbody/tr[1]/td/div/div/div/div/div[1]/div/div/div/div[4]/div[1]/span').get_attribute(
            'innerText')
    except:
        Coord = "Нет данных"

    try:
        driver.find_element_by_id("mts-bs-shortinfo").click()
        time.sleep(3)
        Additional = driver.find_element_by_xpath(
            '//*[@id="comments"]/div/div[2]/div/div/div/div[2]').get_attribute(
            'innerText')
    except:
        Additional = "Нет данных"

    try:
        Podrazdelenie = driver.find_element_by_xpath(
            '/html/body/table/tbody/tr/td[2]/div[1]/table/tbody/tr[2]/td/div[2]/p/span').get_attribute(
            'innerText')
    except:
        Podrazdelenie = "Нет данных"


    return Address,AMS,Coord,Additional,Podrazdelenie


def get_address_PL(driver,PL,to_name):
    BSID = getCedarBSID(driver,PL,to_name)
    driver.get('https://cedar.mts.ru/mrnew/bs/?bsId=%s' % BSID)

    # Адрес А
    Address = driver.find_element_by_xpath(
        '/html/body/table/tbody/tr/td[2]/div[1]/div[2]/table/tbody/tr[1]/td/div/div/div/div/div[1]/div/div/div/div[2]/div[1]/span').get_attribute(
        'innerText')
    try:
        AMS = driver.find_element_by_xpath(
            '//*[@id="mrs-bs-ams"]/div/div[2]/div/div/div/div[2]/div/span').get_attribute(
            'innerText')
    except:
        AMS = "Нет данных по опоре"
    return Address,AMS

def select_RRL(driver,PL):
    time.sleep(2)
    nioss_el = driver.find_elements_by_xpath(
        '/html/body/table/tbody/tr/td[2]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]')
    nioss_el_rrl = driver.find_elements_by_xpath(
        '/html/body/table/tbody/tr/td[2]/div/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr') #/html/body/table/tbody/tr/td[2]/div[7]/table/tbody/tr[3]/td/table/tbody/tr[6]/td[4]/table/tbody/tr/td[1]/a/div
    for n, el in enumerate(nioss_el):
        phase = el.get_attribute('innerText')
        rrl_name = nioss_el_rrl[n].get_attribute('innerText')
        if PL[3:].rjust(5,"0") in rrl_name and phase in ['50', '100']:
            nioss_el_rrl[n].click()
            time.sleep(2)
            return 1
    for n, el in enumerate(nioss_el):
        phase = el.get_attribute('innerText')
        rrl_name = nioss_el_rrl[n].get_attribute('innerText')
        if PL[3:].rjust(5,"0") in rrl_name and phase in ['0']:
            nioss_el_rrl[n].click()
            time.sleep(2)
            return 1
    for n, el in enumerate(nioss_el):
        phase = el.get_attribute('innerText')
        rrl_name = nioss_el_rrl[n].get_attribute('innerText')
        if PL[3:].rjust(5, "0") in rrl_name and phase in ['300','350']:
            nioss_el_rrl[n].click()
            time.sleep(2)
            return 1
    return 0

def getParam_BS(mail_body):
    List_BS=[]
    List_strings_proto =mail_body.split('\n')
    for string in List_strings_proto:
        Sub_string_list=string.split('\r')
        for sub_string in Sub_string_list:
            if len(sub_string)>4:
                try:#рассчитываем, что минимум должно быть 5 символов в номере БС, пробел и один символ описания
                    BS_A=string.strip(' _-').replace('_','').replace('-','').replace('\r','').replace('\n','').replace(' ','')
                    BSR_A = ""
                    for n in BS_A[::-1]:
                        if n.isdigit():
                            BSR_A = n + BSR_A
                    if 8>len(BSR_A)>4:
                        List_BS.append(BSR_A[:2]+"_"+BSR_A[2:])
                except:
                    continue
    return List_BS


def getParam(mail_body):
    List_BS_A_BS_Z=[]
    List_strings_proto =mail_body.split('\n')
    for string in List_strings_proto:
        Sub_string_list=string.split('\r')
        for sub_string in Sub_string_list:
            if len(sub_string)>8:
                try:#рассчитываем, что минимум должно быть 5 символов в номере БС, пробел и один символ описания
                    BS_A=string.split('-')[0].strip(' _-').replace('_','').replace('-','').replace('\r','').replace('\n','').replace(' ','')
                    BS_Z = string.split('-')[1].strip(' _-').replace('_', '').replace('-', '').replace('\r','').replace('\n','').replace(' ','')
                    BSR_A = ""
                    BSR_Z = ""
                    for n in BS_A[::-1]:
                        if n.isdigit():
                            BSR_A = n + BSR_A
                    for n in BS_Z[::-1]:
                        if n.isdigit():
                            BSR_Z = n + BSR_Z
                    if BS_A==BSR_A and BS_Z==BSR_Z:
                        List_BS_A_BS_Z.append([BS_A[:2]+"_"+BS_A[2:],BS_Z[:2]+"_"+BS_Z[2:]])
                except:
                    continue
    return List_BS_A_BS_Z

def save_obj(obj, name ):
    pathname = 'C:\\Users\\yvsobol3\\PycharmProjects\\Test'
    if not os.path.exists(pathname+'\\obj'):
        os.mkdir('obj')
    with open(pathname+'\\obj\\'+ name + '.pkl', 'wb') as f:
        pickle.dump(obj, f, pickle.HIGHEST_PROTOCOL)

def load_obj(name ):
    pathname = 'C:\\Users\\yvsobol3\\PycharmProjects\\Test'
    directory  = os.path.dirname(os.path.realpath(__file__))
    with open(directory+'\\obj\\' + name + '.pkl', 'rb') as f:
        return pickle.load(f)


if __name__ == "__main__":

    org_podryad = {'Телекомсетьстрой (ИНН 5032254494)': 'ООО "Телекомсетьстрой"',
                   'Телус (ИНН 3528177788)': 'ООО "Телус"',
                   'Динат (ИНН 7804397430)': 'ООО "Динат"',
                   'СОЛАРФОН (ИНН 5024098234)': 'ООО "СОЛАРФОН"',
                   'ИНСОЛ ТЕЛЕКОМ (ИНН 7718161470)': 'ООО "Инсол Телеком"',
                   'ООО «СкайТехнолоджиСервис»': 'ООО «СкайТехнолоджиСервис»',
                    'МПО КЛАССИКА (ИНН 7726555460)': 'МПО КЛАССИКА',
                   'Центр-ТЕХОСТРОЙПРОЕКТ (ИНН 0522015457)':'ООО "Центр-ТЕХОСТРОЙПРОЕКТ"',
                   'РИК ООО': 'ООО "РИК"',
                   'ООО ПСК "Юнистрой"':'ООО ПСК "Юнистрой"'}
    lico_podryad = {'Телекомсетьстрой (ИНН 5032254494)': 'Куляпин И.А.',
                    'Телус (ИНН 3528177788)': 'Малышев А.А.',
                    'Динат (ИНН 7804397430)': 'Кокшаров А.Л.',
                    'СОЛАРФОН (ИНН 5024098234)': 'Малкин И.Ю.',
                    'ИНСОЛ ТЕЛЕКОМ (ИНН 7718161470)': 'Путорайкин А.',
                    'ООО «СкайТехнолоджиСервис»': 'Виноградов А.И.',
                    'МПО КЛАССИКА (ИНН 7726555460)': 'Акулов С.В.',
                    'Центр-ТЕХОСТРОЙПРОЕКТ (ИНН 0522015457)': 'Рамазанов З.А.',
                    'РИК ООО': 'Морозов А.М.',
                    'ООО ПСК "Юнистрой"':'Ворчик Ю.С.'}
    SRO_PIR = {'Телекомсетьстрой (ИНН 5032254494)': 'СРО-П-168-22112011',
               'Телус (ИНН 3528177788)': 'СРО № 10207',
               'Динат (ИНН 7804397430)': '№0160-2016-7804397430-05 от 23 сентября 2016г.',
               'СОЛАРФОН (ИНН 5024098234)': 'СРО-П-156-06072010 от 16.11.2018',
               'ИНСОЛ ТЕЛЕКОМ (ИНН 7718161470)': '№ СРО-П-011-16072009 от 28.09.2017г.',
               'ООО «СкайТехнолоджиСервис»': '№ 0253.01-2016-5260325169-П-107 от 03.11.2016',
               'МПО КЛАССИКА (ИНН 7726555460)': 'СРО П-‪011-16072009 от 11.07.2019г',
               'Центр-ТЕХОСТРОЙПРОЕКТ (ИНН 0522015457)': '№ СРО-П-145-04032010 от 03.02.2012 г.',
               'РИК ООО': '№ СРО-П-019-26082009 от 01 июля 2021',
               'ООО ПСК "Юнистрой"':'от 19 июня 2012г.  № 0336.04-2010-61620561-03-П-033'}
    SRO_SMR = {'Телекомсетьстрой (ИНН 5032254494)': 'СРО-С-230-07092010',
               'Телус (ИНН 3528177788)': 'СРО № 6449',
               'Динат (ИНН 7804397430)': '№0564-2015-7804397430-С-010 от 28 октября 2015г.',
               'СОЛАРФОН (ИНН 5024098234)': 'СРО-С-179-20012010 от 15.11.2018г.',
               'ИНСОЛ ТЕЛЕКОМ (ИНН 7718161470)': '№ СРО-С-244-13042012 от 28.05.2015г.',
               'ООО «СкайТехнолоджиСервис»': '№ С-033-52-0464-111016 от 11.10.2016',
                'МПО КЛАССИКА (ИНН 7726555460)': '0115.08-2015-7726555460-С-238 от 26 июня 2015г.',
               'Центр-ТЕХОСТРОЙПРОЕКТ (ИНН 0522015457)': '№ СРО-С-242-13022012 от 03.02.2013 г.',
               'РИК ООО': '№ СРО-С-219-21042010 от 05 июля 2021',
               'ООО ПСК "Юнистрой"':'от 03 июля 2012г.  №1167.04-2010-6162056103-С-031'}

    text=''
    TZ_path_List=[]
    # addMBHlist()
    # sys.argv=load_obj('argv_cedar')
    save_obj(sys.argv, 'argv_cedar')
    try:
        performer,G,initiator=Remedy_Web_1_0.get_performer_by_mail(sys.argv[2],'TN')
        List_BS_A_BS_Z = getParam(sys.argv[1])
    except:
        reply_send(sys.argv[2], '', 'Автоматические исходные данные','Доступ запрещен', "")
        sys.exit()
    print(List_BS_A_BS_Z)

    options = webdriver.ChromeOptions()
    # options.add_argument('headless') #если хотим запустить chrome недивимкой
    options.add_argument("window-size=1920x1080")
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(options=options)
    Credentials = Remedy_Web_1_0.load_obj('Cred')
    login(driver, Credentials['login'], Credentials['pass'])

    for BS_A,BS_Z in List_BS_A_BS_Z:
        try:
            Address_Z,AMS_Z = get_address_PL(driver,BS_Z,sys.argv[2])
            Address_A,AMS_A = get_address_PL(driver,BS_A,sys.argv[2])
            select_RRL(driver, BS_Z)


            # проверяем на реверс A-Z
            revers=False
            nioss_el = driver.find_elements_by_id("RealNumber")
            if nioss_el:
                BS_RRL_A = nioss_el[0].text
                if BS_A[4:].rjust(4,"0")!=BS_RRL_A[-4:]:
                    revers=True

            # Азимут со стороны А
            try:
                AZ_A = driver.find_element_by_xpath(
                    '/html/body/table/tbody/tr/td[2]/div[1]/table/tbody/tr[5]/td[1]/table/tbody/tr/td[2]').get_attribute(
                    'innerText')
            except:
                text = text + ("\nПролет " + BS_Z + " <> " + BS_A + " найти не удалось!"+"\n\n\n")
                continue




            # Азимут со стороны Z
            AZ_Z = driver.find_element_by_xpath(
                '/html/body/table/tbody/tr/td[2]/div[1]/table/tbody/tr[5]/td[2]/table/tbody/tr/td[2]').get_attribute(
                'innerText')

            Dist=driver.find_element_by_xpath(
                '/html/body/table/tbody/tr/td[2]/div[1]/table/tbody/tr[5]/td[1]/table/tbody/tr/td[3]').get_attribute(
                'innerText')

            Coord_A=driver.find_element_by_xpath(
                '/html/body/table/tbody/tr/td[2]/div[1]/table/tbody/tr[5]/td[1]/table/tbody/tr/td[1]').get_attribute(
                'innerText')

            Coord_Z = driver.find_element_by_xpath(
                '/html/body/table/tbody/tr/td[2]/div[1]/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]').get_attribute(
                'innerText')

            Band_text = driver.find_element_by_name('range').get_attribute('value')
            nioss_el = driver.find_elements_by_xpath(
                '/html/body/table/tbody/tr/td[2]/div[1]/table/tbody/tr[7]/td[1]/table/tbody/tr[18]/td/select/option')
            for el in nioss_el:
                if el.get_attribute('value') == Band_text:
                    Band_text = el.get_attribute('text')
                    break

            SubBand_text = driver.find_element_by_name('pd').get_attribute('value')
            nioss_el = driver.find_elements_by_xpath(
                '/html/body/table/tbody/tr/td[2]/div[1]/table/tbody/tr[7]/td[1]/table/tbody/tr[19]/td/select/option')
            for el in nioss_el:
                if el.get_attribute('value') == SubBand_text:
                    SubBand_text = el.get_attribute('text')
                    break

            TX_Power = driver.find_element_by_name('power').get_attribute('value')
            RX_Level = driver.find_element_by_name('calc_level').get_attribute('value')
            Rezerv=driver.find_element_by_name('rezerv').get_attribute('value')
            Polarisation=driver.find_element_by_name('polarization').get_attribute('value')

            BandWidth=driver.find_element_by_name('bandWidth').get_attribute('value')
            nioss_el=driver.find_elements_by_xpath(
                '/html/body/table/tbody/tr/td[2]/div[1]/table/tbody/tr[8]/td/table[2]/tbody/tr[2]/td[2]/select/option')
            for el in nioss_el:
                if el.get_attribute('value') == BandWidth:
                    BandWidth=el.get_attribute('text')
                    break

            EquipmentA=driver.find_element_by_name('aTrans').get_attribute('value')
            nioss_el=driver.find_elements_by_xpath(
                '//*[@id="rrl"]/tbody/tr[1]/td/select/option')
            for el in nioss_el:
                if el.get_attribute('value') == EquipmentA:
                    EquipmentA=el.get_attribute('text')
                    break

            EquipmentZ=driver.find_element_by_name('bTrans').get_attribute('value')
            nioss_el=driver.find_elements_by_xpath(
                '//*[@id="rrl"]/tbody/tr[1]/td/select/option')
            for el in nioss_el:
                if el.get_attribute('value') == EquipmentZ:
                    EquipmentZ=el.get_attribute('text')
                    break

            Ant_A = driver.find_element_by_name('aArrlType1').get_attribute('value')
            nioss_el = driver.find_elements_by_xpath(
                '//*[@id="rrl"]/tbody/tr[6]/td[1]/select/option')
            for el in nioss_el:
                if el.get_attribute('value') == Ant_A:
                    Ant_A = el.get_attribute('text')
                    break

            Ant_Z = driver.find_element_by_name('bArrlType1').get_attribute('value')
            nioss_el = driver.find_elements_by_xpath(
                '//*[@id="rrl"]/tbody/tr[6]/td[1]/select/option')
            for el in nioss_el:
                if el.get_attribute('value') == Ant_Z:
                    Ant_Z = el.get_attribute('text')
                    break

            podryad = driver.find_element_by_xpath("//*[contains(@id,'select2-orgId')]").get_attribute('innerText')
            # nioss_el = driver.find_elements_by_xpath(
            #     '/html/body/table/tbody/tr/td[2]/div[1]/table/tbody/tr[7]/td[2]/table/tbody/tr[11]/td/select/option')
            # for el in nioss_el:
            #     if el.get_attribute('value') == podryad:
            #         podryad = el.get_attribute('text')
            #         break

            TX_freq = driver.find_element_by_name('aFTx1').get_attribute('value')
            RX_freq = driver.find_element_by_name('bFTx1').get_attribute('value')
            Suprvisor=driver.find_element_by_name('curator').get_attribute('value')
            Add_text_1=driver.find_element_by_name('notes').get_attribute('value')
            Add_text_2=driver.find_element_by_name('planNotes').get_attribute('value')
            H_A = driver.find_element_by_name('aH').get_attribute('value')
            H_Z = driver.find_element_by_name('bH').get_attribute('value')



            try:
                org_podryad_text = org_podryad[podryad]
                lico_podryad_text = lico_podryad[podryad]
                SRO_PIR_text = SRO_PIR[podryad]
                SRO_SMR_text = SRO_SMR[podryad]
            except:
                org_podryad_text = podryad
                lico_podryad_text = podryad
                SRO_PIR_text = podryad
                SRO_SMR_text = podryad
            try:
                docTZ = DocxTemplate('c:\\Users\\yvsobol3\\PycharmProjects\\Test\\Шаблон ТЗ.docx')
                if revers:
                    context = {'year': datetime.datetime.now().year,
                           'PL_A': BS_Z,
                           'PL_Z': BS_A,
                           'Address_A': Address_Z,
                           'Address_Z': Address_A,
                           'Type_eqip_A':EquipmentA,
                           'Type_eqip_Z': EquipmentZ,
                           'band':Band_text,
                           'diam_A':Ant_A,
                           'diam_Z' :Ant_Z,
                           "Az1":AZ_A,
                           "Az2": AZ_Z,
                            'rezerv':Rezerv,
                           'region': Address_A.split(',')[0],
                           'H1':H_A,
                           'H2':H_Z,
                           'org_podryad':org_podryad_text,
                           'lico_podryad': lico_podryad_text,
                           'SRO_PIR': SRO_PIR_text,
                           'SRO_SMR': SRO_SMR_text,
                           'kurator':Suprvisor,
                               'AMS_A': AMS_Z,
                               'AMS_Z': AMS_A
                           }
                else:
                    context = {'year': datetime.datetime.now().year,
                               'PL_A': BS_A,
                               'PL_Z': BS_Z,
                               'Address_A': Address_A,
                               'Address_Z': Address_Z,
                               'Type_eqip_A': EquipmentA,
                               'Type_eqip_Z': EquipmentZ,
                               'band': Band_text,
                               'diam_A': Ant_A,
                               'diam_Z': Ant_Z,
                               "Az1": AZ_A,
                               "Az2": AZ_Z,
                               'rezerv': Rezerv,
                               'region': Address_A.split(',')[0],
                               'H1': H_A,
                               'H2': H_Z,
                               'org_podryad': org_podryad_text,
                               'lico_podryad': lico_podryad_text,
                               'SRO_PIR': SRO_PIR_text,
                               'SRO_SMR': SRO_SMR_text,
                               'kurator': Suprvisor,
                               'AMS_A': AMS_A,
                               'AMS_Z': AMS_Z
                               }
                docTZ.render(context)
                TZ_path = 'C:\\ID\\TZ Auto\\Техническое задание на ПИР по РРЛ %s-%s.docx' % (BS_A, BS_Z)
                docTZ.save(TZ_path)

                word = win32com.client.Dispatch('Word.Application')

                TZ_path_pdf = 'C:\\ID\\TZ Auto\\Техническое задание на ПИР по РРЛ %s-%s.pdf' % (BS_A, BS_Z)
                docPDF=word.Documents.Open(TZ_path)
                docPDF.SaveAs(TZ_path_pdf, FileFormat=17)
                TZ_path_List.append(TZ_path)
                TZ_path_List.append(TZ_path_pdf)
                docPDF.Close()
                word.Quit()
            except:
               continue


            # text =text+("\nВремя выполнения программы: %s секунд" % (time.time() - start_time))
            if revers:
                text = text + ("\nПролет " + BS_Z + " <> " + BS_A )
            else:
                text = text + ("\nПролет " + BS_A + " <> " + BS_Z )
            text=text+("\nОборудование А: %s, Оборудование Z: %s" % (EquipmentA, EquipmentZ))
            if revers:
                text=text+("\nАдрес А: %s" % Address_Z+ "\nАдрес Z: %s" % Address_A)
                text = text + ("\nАМС А: %s" % (AMS_Z) + "\nАМС Z: %s" % (AMS_A))
            else:
                text = text + ("\nАдрес А: %s" % Address_A + "\nАдрес Z: %s" % Address_Z)
                text = text + ("\nАМС А: %s"% (AMS_A) + "\nАМС Z: %s" % (AMS_Z))



            text=text+("\nГеография пролета: %s <> %s , %s" % (AZ_A, AZ_Z, Dist)+ "\nСайт А: %s; Сайт Z: %s" % (Coord_A, Coord_Z))
            text=text+("\nДиапазон: %s, поддиапазон: %s" % (Band_text, SubBand_text))
            text = text + ("\nЧастота ТХ А: %s, Частота ТХ Z: %s" % (TX_freq, RX_freq))
            text=text+("\nПоляризация: %s, Ширина спектра: %s, Резервирование: %s" % (Polarisation, BandWidth, Rezerv)+
                  "\nАнтенна А: %s, Антенна Z: %s" % (Ant_A, Ant_Z))
            text=text+("\nМощность ТХ: %s, Уровень RX    (расчет): %s" % (TX_Power, RX_Level)+
                  "\nДополнительно: %s,\nСтатус работ: %s" % (Add_text_1, Add_text_2)+ "\nКуратор: %s" % Suprvisor)
            text = text + "\n\n\n"
        except:
            text = text + ("\nПролет " + BS_Z + " <> " + BS_A + " не отработан! Проверьте корретность заведенных данных.\n\n\n")
        print(text)
    text = "Запрос отработан успешно!"+text
    reply_send(sys.argv[2], performer, 'Автоматические исходные данные', text,TZ_path_List)
    driver.quit()
    sys.exit()
