
import NIOSS
import time
import datetime
import xlrd
import sys
import Remedy_Web_1_0
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver



def RRL_add_AGOS(driver, HREF):
    now = datetime.datetime.now()
    # try:
    # открываем ближнюю сторону РРС
    driver.get(HREF[0])
    # считываем айпи адрес, если он есть
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9136474451313754458')))
    time.sleep(1.5)
    try:
        driver.find_element_by_link_text("Документы").click()
    except:
        driver.find_element_by_link_text("Больше").click()
        driver.find_element_by_link_text("Документы").click()
    driver.find_element_by_link_text("Новый АГОС").click()
        #     необходимо проверить есть ли уже заведенный агос
    #  заполняем дату
            # отправляем значения в браузер
    nioss_el = driver.find_elements_by_xpath("//input[@type='text']")
    nioss_el[2].send_keys(now.day)
    nioss_el[3].send_keys(now.month)
    nioss_el[4].send_keys(now.year)
    driver.find_element_by_link_text("Создать").click()
    try:
        driver.find_element_by_id("pcUpdate").click()
    except:
        print("С одной стороны АГОС существует.")

    driver.get(HREF[2])
    # считываем айпи адрес, если он есть
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9136474451313754458')))
    time.sleep(1.5)
    try:
        driver.find_element_by_link_text("Документы").click()
    except:
        driver.find_element_by_link_text("Больше").click()
        driver.find_element_by_link_text("Документы").click()
    driver.find_element_by_link_text("Новый АГОС").click()
    #     необходимо проверить есть ли уже заведенный агос
    #  заполняем дату
    # отправляем значения в браузер
    nioss_el = driver.find_elements_by_xpath("//input[@type='text']")
    nioss_el[2].send_keys(now.day)
    nioss_el[3].send_keys(now.month)
    nioss_el[4].send_keys(now.year)
    driver.find_element_by_link_text("Создать").click()
    try:
        driver.find_element_by_id("pcUpdate").click()
    except:
        print("С другой стороны АГОС существует.")

        # функция для добавления устройства в ниосс, принимает обджект айди и тип узла.

def MBH_Demont_device(driver, ID):
    now = datetime.datetime.now()
    # try:
    # открываем файл ексель, который должен содерждать данные о типе оборудования
    # открываем ближнюю сторону РРС
    driver.get("https://nioss/ncobject.jsp?id=" + ID)
    time.sleep(1.5)
    if 'эксплуата' in driver.find_element_by_xpath('//*[@id="vv_9130962698813156982"]/span/span/span').get_attribute('innerText'):
        print('необходимо отвязать сетевые элементы от узла https://nioss/ncobject.jsp?id=' + ID)

        driver.find_element_by_link_text("Редактировать").click()
        NIOSS.set_value(driver, 'id__4_6040759067013572906_'+ID+'_0','23')
        NIOSS.set_value(driver, 'id__4_6040759067013572906_' + ID + '_3', '12')
        NIOSS.set_value(driver, 'id__4_6040759067013572906_' + ID + '_6', '2019')
        driver.find_element_by_link_text("Обновить").click()
        #     необходимо проверить есть ли уже заведенный агос
        #  заполняем дату
        # отправляем значения в браузер


def MBH_Demont_device_list():
    rb = xlrd.open_workbook('C:/Users/yvsobol3/PycharmProjects/Test/MBH_delete_device.xlsx')
    sheet = rb.sheet_by_index(0)
    start_time = time.time()

    options = webdriver.ChromeOptions()
    # options.add_argument('headless') #если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    driver = webdriver.Chrome(options=options)
    Credentials = Remedy_Web_1_0.load_obj('Cred')

    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        if row[0]:
            try:
                MBH_Demont_device(driver, row[0])
                # print('Строка ' + str(rownum) +  ' отработана')
            except:
                # print('Строка ' + str(rownum)+ ' не отработана')
                driver.close()
                driver = webdriver.Chrome(options=options)
                NIOSS.NiossLogin(driver, Credentials['login'], Credentials['pass'])


# функция, заносящая узел в НИОСС
def MBH_add_to_NIOSS(driver,PL, MBH_name, IP,proekt,task, type_device):
    now = datetime.datetime.now()
    driver.get("https://nioss/common/search.jsp?o=1000&object=9130921577213339143&explorer_mode=disable")
    # перемещаемся на страницу поиска площадки
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, '_r_1')))
    select = Select(driver.find_element_by_name('_r_1'))
    select.select_by_visible_text('равно')
    driver.find_element_by_id("_v_1").clear()
    driver.find_element_by_id("_v_1").send_keys(PL)
    driver.find_element_by_id("_v_1").send_keys(Keys.ENTER)

    driver.find_element_by_link_text(PL).click()
    # получаем список СЭ и проверяем на совпадение по имени
    time.sleep(4)
    nioss_el = driver.find_elements_by_xpath('// *[ contains (@ id,"jx_cell_")]')
    innerTextList = []
    for el in nioss_el:
        if MBH_name==el.get_attribute("innerText"):
            print("Элемент с данным именем присутствует в НИОСС")
            return

    # // *[ @ id = "jx_cell_t3100965357013861641_0_t_9134852835013305937_-1"] / a
    # jx_cell_t3100965357013861641_0_t_9135621006013379805_-1 > a
    driver.find_element_by_link_text("Новый Элемент сети").click()

    # перемещаемся на страницу создания СЭ
    # Создаем MBH

    try:
        alert = driver.switch_to_alert()
        alert.accept()
        driver.switch_to_default_content()
    except:
        print("Уведомления не было!")
    driver.find_elements_by_id('combo_arrowlookin')[0].click()
    driver.find_element_by_xpath('//*[@id="lookin_nodeIcon36"]').click()
    driver.find_element_by_xpath('//*[@id="lookin_nodeIcon60"]').click()
    driver.find_element_by_link_text("Узел MBH").click()
    NIOSS.sel_value(driver, 'id__7_9139913988213651937', 'Мобильная сеть')
    NIOSS.set_value(driver, 'id__9_6040759067013572907_input', 'Tellabs')
    NIOSS.set_value(driver, 'id__9_9161659439113836003_input', proekt)
    driver.find_element_by_xpath('//*[@id="theform"]/table/tbody/tr/td/a[1]/img').click()  # создать
    time.sleep(2)
    ID=NIOSS.getObgectID(driver.current_url)
    driver.find_element_by_id('pcEdit').click()
    try:
        IPlist = IP.split('.')
        if len(IPlist)==4:
            NIOSS.set_value(driver, 'id__5_9136474451313754458_' + ID + '_0', IPlist[0])
            NIOSS.set_value(driver, 'id__5_9136474451313754458_' + ID + '_3', IPlist[1])
            NIOSS.set_value(driver, 'id__5_9136474451313754458_' + ID + '_6', IPlist[2])
            NIOSS.set_value_without_enter(driver, 'id__5_9136474451313754458_' + ID + '_9', IPlist[3])
        # driver.find_element_by_id('pcEdit').click()
        NIOSS.set_value_without_enter(driver, 'id__0_7040238585013737267_' + ID , MBH_name)
        NIOSS.set_value_without_enter(driver, 'id__0_9062885501013447728_' + ID, proekt)
        NIOSS.set_value(driver, 'id__0_9062885501013447727_' + ID, task)
        driver.find_element_by_id('pcUpdate').click()
        print("Заведен " + MBH_name)
    except Exception as ex:
        print('Ошибка: ')
        print(ex)
    MBH_add_device(driver, ID, MBH_name, type_device)

    # функция для добавления устройства в ниосс, принимает обджект айди и тип узла.
def MBH_add_device(driver, ID,MBH_name,type_device):
    now = datetime.datetime.now()
    # try:
    # открываем файл ексель, который должен содерждать данные о типе оборудования
    # открываем ближнюю сторону РРС
    if '8615' in type_device:
        type_device_nioss="8615 Sm"
    elif '8609' in type_device:
        type_device_nioss="8609 Sm"
    elif '8603' in type_device:
        if ('A' in type_device) or ('А' in type_device):
            type_device_nioss="8603-A"
        else:
            type_device_nioss = "8603-B"
    elif '8611' in type_device:
        type_device_nioss="8611 Sm"
    elif '8625' in type_device:
        type_device_nioss="Tellabs 8625"
    elif '8665' in type_device:
        type_device_nioss="Tellabs 8665"
    else:
        type_device_nioss = ""

    driver.get("https://nioss/ncobject.jsp?id=" + ID)
    time.sleep(1.5)
    try:
        driver.find_element_by_link_text("Устройства").click()
    except:
        driver.find_element_by_link_text("Больше").click()
        driver.find_element_by_link_text("Устройства").click()
    driver.find_element_by_link_text("Добавить новое устройство").click()
        #     необходимо проверить есть ли уже заведенный агос
    #  заполняем дату
            # отправляем значения в браузер
    nioss_el = driver.find_elements_by_xpath('//*[@id="main_window"]/tbody/tr[4]/td[2]/input')

    nioss_el[0].send_keys(type_device_nioss)
    nioss_el = driver.find_elements_by_xpath('//*[@id="main_window"]/tbody/tr[7]/td/div[2]/div/a')
    nioss_el[0].click()
    driver.find_element_by_link_text("Создать").click()


    # функция для добавления устройства в ниосс по списку,открывает файл в локальной папке.
def MBH_add_device_list():
    Credentials = {'login': 'yvsobol3', 'pass': 'Cj,jk123'}
    rb = xlrd.open_workbook('C:/Users/yvsobol3/PycharmProjects/Test/MBH_add_device.xlsx')
    sheet = rb.sheet_by_index(0)
    start_time = time.time()

    options = webdriver.ChromeOptions()
    # options.add_argument('headless') #если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    driver = webdriver.Chrome(options=options)
    NIOSS.NiossLogin(driver, Credentials['login'], Credentials['pass'])

    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        if row[0] and row[1] and row[2]:
            try:
                MBH_add_device(driver, row[0], row[1],row[2])
                print ('Строка ' +str(rownum)+ ' '+row[1]+' отработана')
            except:
                print('Строка ' + str(rownum) + ' '+row[1]+' не отработана')
                driver.close()
                driver = webdriver.Chrome(options=options)
                NIOSS.NiossLogin(driver, Credentials['login'], Credentials['pass'])


    # функция для создания РРЛ из двух РРС.
def RRL_create(driver, RRS_A,RRS_Z,Site_A,Site_Z):
    # try:
    # открываем файл ексель, который должен содерждать данные о типе оборудования
    # открываем ближнюю сторону РРС
    driver.get("https://nioss/ncobject.jsp?id=6032252332013374327&site=9130808606313319033")
    time.sleep(1.5)
    # заносим сайт А -Z
    driver.find_element_by_link_text("Новое Радиорелейное соединение").click()
    nioss_el = driver.find_elements_by_xpath('//*[@id="theform"]/table[2]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/input')
    nioss_el[0].send_keys(Site_A)
    nioss_el = driver.find_elements_by_xpath('//*[@id="theform"]/table[2]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/a/img')
    nioss_el[0].click()


    nioss_el = driver.find_elements_by_xpath('//*[@id="theform"]/table[2]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/select/option')
    for el in nioss_el:
        attr=el.get_attribute("innerText")
        attr=attr.split('->')[-1]
        if attr == Site_A:
            el.click()
            break




    nioss_el = driver.find_elements_by_xpath('//*[@id="theform"]/table[2]/tbody/tr[1]/td[2]/table/tbody/tr[1]/td[2]/input')
    nioss_el[0].send_keys(Site_Z)
    nioss_el = driver.find_elements_by_xpath('//*[@id="theform"]/table[2]/tbody/tr[1]/td[2]/table/tbody/tr[1]/td[2]/a/img')
    nioss_el[0].click()

    nioss_el = driver.find_elements_by_xpath(
        '//*[@id="theform"]/table[2]/tbody/tr[1]/td[2]/table/tbody/tr[1]/td[2]/select/option')
    for el in nioss_el:
        attr = el.get_attribute("innerText")
        attr = attr.split('->')[-1]
        if attr == Site_Z:
            el.click()
            break




    #  заносим РРС
    nioss_el = driver.find_elements_by_xpath(
        '//*[@id="theform"]/table[2]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td[2]/input')
    nioss_el[0].send_keys(RRS_A)
    nioss_el = driver.find_elements_by_xpath(
        '//*[@id="theform"]/table[2]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td[2]/a/img')
    nioss_el[0].click()
    nioss_el = driver.find_elements_by_xpath(
        '//*[@id="theform"]/table[2]/tbody/tr[1]/td[2]/table/tbody/tr[2]/td[2]/input')
    nioss_el[0].send_keys(RRS_Z)
    nioss_el = driver.find_elements_by_xpath(
        '//*[@id="theform"]/table[2]/tbody/tr[1]/td[2]/table/tbody/tr[2]/td[2]/a/img')
    nioss_el[0].click()
    #заносим какую-нибудь антенну.
    # nioss_el = driver.find_element_by_xpath(
    #     '//*[@id="id__9_3100862885013916366_input"]')
    # nioss_el.send_keys('Andrew 13/1.8 (VHP6-130)')
    # time.sleep(1)
    # nioss_el.send_keys(Keys.ENTER)
    # time.sleep(1)
    # nioss_el = driver.find_element_by_xpath(
    #     '//*[@id="id__9_3100862885013916367_input"]')
    # nioss_el.send_keys('Andrew 13/1.8 (VHP6-130)')
    # time.sleep(1)
    # nioss_el.send_keys(Keys.ENTER)
    # time.sleep(1)
    nioss_el = driver.find_elements_by_xpath(
        '//*[@id="theform"]/table[2]/tbody/tr[2]/td/a[1]/img')
    nioss_el[0].click()
    time.sleep(5)
    return driver.current_url


def RRL_create_list():
    Credentials = {'login': 'yvsobol3', 'pass': 'Cj,jk123'}
    rb = xlrd.open_workbook('C:/Development/Макросы/Генерирование РРС из Кедра.xlsm')
    sheet = rb.sheet_by_index(3)
    start_time = time.time()

    options = webdriver.ChromeOptions()
    # options.add_argument('headless') #если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    driver = webdriver.Chrome(options=options)
    NIOSS.NiossLogin(driver, Credentials['login'], Credentials['pass'])

    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        if row[0:3]:
            # try:
            RRL_create(driver, row[2], row[3],row[0],row[1])
            print ('Строка ' +str(rownum)+ ' '+row[2]+' отработана')
            # except:
            #     print('Строка ' + str(rownum) + ' '+row[2]+' не отработана')
            #     driver.close()
            #     driver = webdriver.Chrome(options=options)
            #     NIOSS.NiossLogin(driver, Credentials['login'], Credentials['pass'])


   # функция для создания РРЛ из двух РРС.
def RRN_connect_RRS(driver, RRS_HREF,RRN_Name):
    # try:
    # открываем файл ексель, который должен содерждать данные о типе оборудования
    # открываем ближнюю сторону РРС
    time.sleep(1.5)
    driver.get(RRS_HREF)

    driver.find_element_by_id('pcEdit').click()
    # заносим сайт А -Z
    time.sleep(1.5)
    print(str(RRS_HREF))
    RRS_ID=NIOSS.getObgectID(str(RRS_HREF),'id=')
     # *[ @ id = "id__9_9133816070113745923_9139670853313398746_input"]
    # driver.find_elements_by_xpath('//*[@id="id__9_9143681853213133160_'+RRS_ID+'_div"]/div/i')[0].click()

    NIOSS.set_value(driver,'id__9_9133816070113745923_'+RRS_ID+'_input',RRN_Name)
    # time.sleep(4)
    # nioss_el = driver.find_elements_by_xpath('//*[@class="refsel_name"]')
    #
    # for el in nioss_el:
    #     attr = el.text
    #
    #     if attr and RRS in attr:
    #         el.click()
    #         break
    time.sleep(15)
    driver.find_element_by_id('theform_update').click()


def EXT_connect_RRS(driver, RRS_HREF,RRN_Name):
    # try:
    # открываем файл ексель, который должен содерждать данные о типе оборудования
    # открываем ближнюю сторону РРС
    time.sleep(1.5)
    driver.get(RRS_HREF)

    driver.find_element_by_id('pcEdit').click()
    # заносим сайт А -Z
    time.sleep(1.5)
    print(str(RRS_HREF))
    RRS_ID=NIOSS.getObgectID(str(RRS_HREF),'id=')
     # *[ @ id = "id__9_9133816070113745923_9139670853313398746_input"]
    # driver.find_elements_by_xpath('//*[@id="id__9_9143681853213133160_'+RRS_ID+'_div"]/div/i')[0].click()

    NIOSS.set_value_without_clear(driver,'id__9_9133816070113745923_'+RRS_ID+'_input',RRN_Name)
    # time.sleep(4)
    # nioss_el = driver.find_elements_by_xpath('//*[@class="refsel_name"]')
    #
    # for el in nioss_el:
    #     attr = el.text
    #
    #     if attr and RRS in attr:
    #         el.click()
    #         break
    time.sleep(15)
    driver.find_element_by_id('theform_update').click()



def check_BS_NIOSS(driver, Site_A):
    time.sleep(3)
    url = 'https://nioss/common/search.jsp?search_mode=search&resultID=9162674715113753096&performed=true&fast_search=no&do_search=yes&ctrl=t4122758596013619618_4122758596013619619&tab=0&return=%2Fstartpage.jsp&currobject=o&system_index_on=false&collapse_but=yes&profile_id=9130921577213339143&object=9130921577213339143&o=1000&explorer_mode=disable&property_ishidden_widget_description=no&property_ishidden_group_-10=no&_r_1=eq&_v_1='+Site_A+'&_r_2=eq&property_ishidden_group_1090440074013323208=yes&_r6022259471013260638=eq&property_ishidden_group_9138844652013066929=yes&_r9138844652013066914=eq&property_ishidden_group_9138644316913567553=yes&_r9138645062513570791=in&_r9138645066113570792=in&_r9138645069513570793=in&_r9138645295113573925=in&_r9138645293813573924=eq&_r6071758452013983195=eq&_r9144751040013465690=in&_r9144751040013465696=in&_r9144752517313465900=in&_r9144752517313465901=eq&_r9144752517313465902=eq&_r9144692932813442513=eq&property_ishidden_widget_scope_options=yes&currproject=default&remember=yes&property_ishidden_widget_templates=yes'
    driver.get(url)
    time.sleep(5)
    # перемещаемся на страницу поиска площадки
    # select = Select(driver.find_element_by_name('_r_1'))
    # select.select_by_visible_text('равно')
    # NIOSS.set_value(driver,"_v_1",Site_A)
    # '//*[@id="t4122361118013615427_9138014405313925172_tbody"]/tr/td[2]/a/span'
    # '//*[@id="t4122361118013615427_9138014405313925172_tbody"]/tr/td[2]/a/span'

    try:
        el=driver.find_element_by_link_text(Site_A)
        a=el.get_attribute('href')
    except:
        a=''
    return a


def RRN_connect_RRS_list():
    Credentials = {'login': 'yvsobol3', 'pass': 'Cj,jk123'}
    rb = xlrd.open_workbook('C:/Development/Макросы/Генерирование РРС из Кедра.xlsm')
    sheet = rb.sheet_by_index(3)
    start_time = time.time()

    options = webdriver.ChromeOptions()
    # options.add_argument('headless') #если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    driver = webdriver.Chrome(options=options)
    NIOSS.NiossLogin(driver, Credentials['login'], Credentials['pass'])

    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        if row[0:1]:
            # try:
                RRN_connect_RRS(driver, row[0], row[1])
                print ('Строка ' +str(rownum)+ ' '+row[1]+' отработана')
            # except:
            #     print('Строка ' + str(rownum) + ' '+row[1]+' не отработана')
            #     driver.close()
            #     driver = webdriver.Chrome(options=options)
            #     NIOSS.NiossLogin(driver, Credentials['login'], Credentials['pass'])


def get_list_rrs_by_bshref(driver,bs_href,NE_ststus_str='Планируется'):
    time.sleep(2)
    rrns=[]
    driver.get(bs_href)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "t3100965357013861641_0_tbody"))
    )
    time.sleep(1)
    names=driver.find_elements_by_xpath('//*[@id="t3100965357013861641_0_tbody"]/tr/td[2]/a/span')
    for i,name in enumerate(names):
            try:
                status = driver.find_element_by_xpath('//*[@id="t3100965357013861641_0_tbody"]/tr['+str(i+1)+']/td[3]/span').get_attribute('innerText')
            except:
                status=''
            try:
                href = driver.find_element_by_xpath('//*[@id="t3100965357013861641_0_tbody"]/tr['+str(i+1)+']/td[2]/a').get_attribute('href')
            except:
                href=''
            if status==NE_ststus_str:


                rrns.append([name.get_attribute('innerText'),status,href])
    return rrns
    '//*[@id="t3100965357013861641_0_tbody"]/tr[84]/td[3]/span'

def get_list_NE_by_bshref(driver,bs_href):

    rrns=[]
    time.sleep(2)
    driver.get(bs_href)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 't3100965357013861641_0_tbody'))
    )
    names=driver.find_elements_by_xpath('//*[@id="t3100965357013861641_0_tbody"]/tr/td[2]/a/span')
    for i,name in enumerate(names):
            try:
                status = driver.find_element_by_xpath('//*[@id="t3100965357013861641_0_tbody"]/tr['+str(i+1)+']/td[3]/span').get_attribute('innerText')
            except:
                status=''
            try:
                typeRRS = driver.find_element_by_xpath('//*[@id="t3100965357013861641_0_tbody"]/tr['+str(i+1)+']/td[6]/a').get_attribute('innerText')
            except:
                typeRRS=''
            try:
                typeRRN = driver.find_element_by_xpath('//*[@id="t3100965357013861641_0_tbody"]/tr['+str(i+1)+']/td[5]/a').get_attribute('innerText')
            except:
                typeRRN=''
            try:
                ip=driver.find_element_by_xpath('//*[@id="t3100965357013861641_0_tbody"]/tr['+str(i+1)+']/td[7]').get_attribute('innerText')
            except:
                ip=''
            try:
                href = driver.find_element_by_xpath('//*[@id="t3100965357013861641_0_tbody"]/tr['+str(i+1)+']/td[2]/a').get_attribute('href')
            except:
                href=''
            rrns.append([name.get_attribute('innerText'),typeRRN,typeRRS,ip,status,href])
    return rrns


def RRL_birth(driver,RRL):
    # try:
    error=False
    error_text=''
    Credentials = {'login': 'yvsobol3', 'pass': 'Cj,jk135'}
    if not RRL.RRS_A_Href:
        RRL.RRS_A_Href, RRL.RRS_A_name=RRS_create(driver, RRL.site_PL_A,RRL.site_PL_Z,RRL.PL_A_Href)
    if not RRL.RRS_Z_Href:
        RRL.RRS_Z_Href, RRL.RRS_Z_name= RRS_create(driver, RRL.site_PL_Z, RRL.site_PL_A,RRL.PL_Z_Href)

    if not RRL.RRL_Href:
        RRL.RRL_Href=RRL_create(driver,RRL.RRS_A_name,RRL.RRS_Z_name,RRL.site_PL_A,RRL.site_PL_Z)

    if not RRL.RRN_A_Href:
        RRL.RRN_A_Href, RRL.RRN_A_name = RRN_create(driver, RRL.PL_A_Href)
    if not RRL.RRN_A_exist_bool:
        RRN_connect_RRS(driver, RRL.RRS_A_Href, RRL.RRN_A_name)
    if not RRL.RRN_Z_Href:
        RRL.RRN_Z_Href, RRL.RRN_Z_name = RRN_create(driver, RRL.PL_Z_Href)
    if not RRL.RRN_Z_exist_bool:
        RRN_connect_RRS(driver, RRL.RRS_Z_Href, RRL.RRN_Z_name)

    if RRL.EXT_A:
        if not RRL.EXT_A_Href:
            RRL.EXT_A_Href, RRL.EXT_A_name = EXT_create(driver, RRL.PL_A_Href)
        if not RRL.EXT_A_exist_bool:
            EXT_connect_RRS(driver, RRL.RRS_A_Href, RRL.EXT_A_name)
    if RRL.EXT_Z:
        if not RRL.EXT_Z_Href:
            RRL.EXT_Z_Href, RRL.EXT_Z_name = EXT_create(driver, RRL.PL_Z_Href)
        if not RRL.EXT_Z_exist_bool:
            EXT_connect_RRS(driver, RRL.RRS_Z_Href, RRL.EXT_Z_name)



    return error,RRL



def RRS_create(driver,Site_A,Site_Z,siteA_HREF):
    time.sleep(2)
    driver.get(siteA_HREF)
    # перемещаемся на страницу поиска площадки

    # получаем список РРС и определяем порядковый номер
    # // *[ @ id = "t3100965357013861641_0_tbody"] / tr[1] / td[2] / a / span
    # // *[ @ id = "t3100965357013861641_0_tbody"] / tr[3] / td[2] / a / span
    nioss_el=driver.find_elements_by_xpath('//*[@id="t3100965357013861641_0_tbody"]/tr/td[2]/a/span')
    innerTextList=[]
    for el in nioss_el:
        innerTextList.append(el.get_attribute("innerText"))
    RRS='RRS_'+Site_A[3:] + '<>' + Site_Z[3:]
    k=1
    while RRS in innerTextList:
        RRS = 'RRS_' + Site_A[3:] + '<>' + Site_Z[3:]+"_"+str(k)
        k=k+1

    # // *[ @ id = "jx_cell_t3100965357013861641_0_t_9134852835013305937_-1"] / a
    # jx_cell_t3100965357013861641_0_t_9135621006013379805_-1 > a
    driver.find_element_by_link_text("Новый Элемент сети").click()

    # перемещаемся на страницу создания СЭ
    # Создаем РРС А


    try:
        alert = driver.switch_to_alert()
        alert.accept()
        driver.switch_to_default_content()
    except:
        print("Уведомления не было!")
    driver.find_elements_by_id('combo_arrowlookin')[0].click()
    driver.find_element_by_link_text("Радиорелейная Станция").click()
    nioss_el = driver.find_element_by_xpath('// *[ @ id = "theform"] / table / tbody / tr[1] / td[2] / input')
    nioss_el.clear()
    nioss_el.send_keys(RRS)

    NIOSS.sel_value(driver, 'id__7_9139913988213651937', 'Мобильная сеть')
    driver.find_element_by_xpath('//*[@id="theform"]/table/tbody/tr/td/a[1]/img').click()#создать
    time.sleep(10)
    return driver.current_url,driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[1]/div[2]/h1/a').text

def RRN_create(driver,site_HREF):
    time.sleep(3)
    driver.get(site_HREF)
    # перемещаемся на страницу поиска площадки
    time.sleep(2)

    driver.find_element_by_link_text("Новый Элемент сети").click()

    # перемещаемся на страницу создания СЭ
    # Создаем РРС А

    driver.find_elements_by_id('combo_arrowlookin')[0].click()
    driver.find_element_by_xpath('// *[ @ id = "lookin_nodeIcon58"]').click()
    driver.find_element_by_link_text("Радио-релейный узел").click()
    NIOSS.sel_value(driver, 'id__7_9139913988213651937', 'Мобильная сеть')
    driver.find_element_by_xpath('//*[@id="theform"]/table/tbody/tr/td/a[1]/img').click()  #создать
    time.sleep(5)
    return driver.current_url,driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[1]/div[2]/h1/a').text

def EXT_create(driver,site_HREF):
    time.sleep(3)
    driver.get(site_HREF)
    # перемещаемся на страницу поиска площадки
    time.sleep(2)

    driver.find_element_by_link_text("Новый Элемент сети").click()

    # перемещаемся на страницу создания СЭ
    # Создаем РРС А

    driver.find_elements_by_id('combo_arrowlookin')[0].click()
    driver.find_element_by_xpath('// *[ @ id = "lookin_nodeIcon58"]').click()
    driver.find_element_by_xpath('// *[ @ id = "lookin_nodeIcon61"]').click()
    driver.find_element_by_link_text("Радио-релейный узел EXT").click()
    NIOSS.sel_value(driver, 'id__7_9139913988213651937', 'Мобильная сеть')
    driver.find_element_by_xpath('//*[@id="theform"]/table/tbody/tr/td/a[1]/img').click()  #создать
    time.sleep(5)
    return driver.current_url,driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[1]/div[2]/h1/a').text

if __name__ == "__main__":
    start_time = time.time()
    options = Remedy_Web_1_0.webdriver.ChromeOptions()
    # options.add_argument('headless')  # если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    options.add_experimental_option('useAutomationExtension', False)
    driver = Remedy_Web_1_0.webdriver.Chrome(options=options)

    for value in sys.argv[1].split("\r")[:-1]:
        # Получаем словарь из Excel
        list_param = value.split("-:")
        list_param.remove('')
        for n, v in enumerate(list_param):
            list_param[n] = v.split("=")
        dicVal = dict(list_param)

        if sys.argv[2] == "New_MBH_NIOSS":
            MBH_add_to_NIOSS(driver,dicVal['PL'],dicVal['MBH'],dicVal['IP'],dicVal['Proekt'],dicVal['Task'],dicVal['Type'])
