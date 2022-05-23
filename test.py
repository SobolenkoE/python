import RRL_Class
import pythonping
import SNMP
import NIOSS
import time
import pickle
import simplekml
from selenium.webdriver.common.keys import Keys
import SNMP
import Remedy_Web_1_0
import Remedy_Web
import Main
import AutoWin
import RRL_Class
from pysnmp.entity.rfc3413.oneliner import cmdgen
from selenium.webdriver import ActionChains
import AutoDO
from docxtpl import DocxTemplate
import datetime
import AutoDO
import Cedar
import win32com.client
import pymysql
import xlrd
import os
import sys
from retry import retry
import urllib.request
from selenium import webdriver




if __name__ == "__main__":

    # urllib.request.urlretrieve('https://cedar.mts.ru/basic/web/reports/report/download?id=238897','file.xlsx')
    # obj={}
    # if not os.path.exists('obj'):
    #     os.mkdir('obj')
    # with open('obj/Cred.pkl', 'wb') as f:
    #     obj['login'], obj['pass'] = 'yvsobol3', 'Cj,jk134'
    #     pickle.dump(obj, f, pickle.HIGHEST_PROTOCOL)


    # options = webdriver.ChromeOptions()
    # options.add_argument('--profile-directory=Default')
    # # options.add_argument('headless')  # если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    # driver = webdriver.Chrome(options=options)
    # Remedy_Web.remedyLogin(driver, 'yvsobol3', 'Cj,jk135')




    pathname = os.path.dirname(sys.argv[0])
    options = webdriver.ChromeOptions()
    # options.add_argument('--allow-profiles-outside-user-dir')
    # options.add_argument('--enable-profile-shortcut-manager')
    # # options.add_argument('user-data-dir=C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user')
    # # options.add_argument('--profile-directory=Profile 1')
    # # options.add_argument('--profile-directory=Default')
    # options.add_argument('--disable-gpu')
    #
    # options.add_argument("--disable-notifications")
    # # if '!' not in sys.argv[3].lower():
    # # options.add_argument('headless') #если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    # options.add_experimental_option('useAutomationExtension', False)

    # while True:
    driver = Remedy_Web_1_0.get_driver(options)
    #     driver.get('http://remedy.msk.mts.ru/')
    #     cookieID=driver.get_cookie('JSESSIONID')
    #     print(cookieID['value'])
    #     time.sleep(5)
    #     driver.delete_all_cookies()
    #     driver.quit()



    Remedy_Web_1_0.remedyLogin(driver)
    # Remedy_Web_1_0.driver_quit(driver)
    # Cedar.addRRLlist()
    # str = pythonping.ping('12.76.13.161', size=40, count=4)

    # BS , describeW, performer, performerG = '23_0615', 'Модернизация пролета', 'Максименко Олег Николаевич', 'ЮГ\Краснодар\ОРС\ТС'


    # Remedy_Web_1_0.remedyLogin(driver)

    # options = webdriver.ChromeOptions()
    #
    # driver = webdriver.Chrome(options=options)
    # Credentials = Remedy_Web_1_0.load_obj('Cred')
    # NIOSS.NiossLogin(driver, Credentials['login'], Credentials['pass'])
    # driver.get("https://nioss/ncobject.jsp?id=6032252332013374327&site=9130808606313319033")
    # time.sleep(1.5)
    # driver.find_element_by_link_text("Новое Радиорелейное соединение").click()
    # time.sleep(1.5)
    # nioss_el = driver.find_element_by_xpath(
    #     '//*[@id="id__9_3100862885013916366_input"]')
    # nioss_el.send_keys('Andrew 13/1.8 (VHP6-130)')
    # time.sleep(1.5)
    # nioss_el.send_keys(Keys.ENTER)
    # nioss_el = driver.find_element_by_xpath(
    #     '//*[@id="id__9_3100862885013916367_input"]')
    # nioss_el.send_keys('Andrew 13/1.8 (VHP6-130)')
    # nioss_el.send_keys(Keys.ENTER)
    # # impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
    # # klassificator = ['02', '04']
    # # addInfo = ''
    # # supervisor = 'Оперативный дежурный ООКУ СРД ФО'
    # # days = 1
    # # typeNE = 'БС'
    # # NE_impact=[]_
    # NE_impact.append(BS)
    # NE_impact.append('23_0368')
    # # в этой секции анализируется текст сообщения для того, чтобы по словам маркерам выбрать классификатор
    # # по умолчанию классификатор "монтаж оборудования"
    #
    #
    # impactS = 'Незначительное/кратковременное влияние на сервис абонентам в зоне действия СЭ.'
    #
    # Remedy_Web_1_0.remedy_set_param(driver, typeW='Планово-профилактическая',
    #                typeS='Данные + Голос',
    #                impactS=impactS,
    #                addInfo=addInfo,
    #                supervisor=supervisor,
    #                initiator='',
    #                BS=BS, describeW=describeW, performer=performer, performerG=performerG,
    #                typeNE=typeNE, klassificator=klassificator,
    #                days=days,NE_impact=NE_impact)
    #
    # print("Время выполнения программы: %s секунд" % (time.time() - start_time))
    #
    # driver.quit()
