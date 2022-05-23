from selenium import webdriver
import pickle
import sys
import os
import json
from NIOSS import set_value
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from tkinter import messagebox
# import Main
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common import exceptions as EX
from selenium.webdriver.common.keys import Keys
from retry import retry
from selenium.common.exceptions import WebDriverException
import win32com.client
import re, os, time, datetime, uuid


def save_obj(obj, name ):
    if not os.path.exists('obj'):
        os.mkdir('obj')
    with open('obj/'+ name + '.pkl', 'wb') as f:
        pickle.dump(obj, f, pickle.HIGHEST_PROTOCOL)

def load_obj(name ):
    directory  = os.path.dirname(os.path.realpath(__file__))
    with open(directory+'\\obj\\' + name + '.pkl', 'rb') as f:
        return pickle.load(f)

def get_driver(options):
    options.add_argument('--allow-profiles-outside-user-dir')
    options.add_argument('--enable-profile-shortcut-manager')
    options.add_argument('user-data-dir=C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user')
    # options.add_argument('--profile-directory=C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user\\Profile 1')
    # options.add_argument('--profile-directory=Default')
    options.add_argument('--disable-gpu')

    # options.add_argument("--disable-notifications")
    # if '!' not in sys.argv[3].lower():
    # options.add_argument('headless') #если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    options.add_experimental_option('useAutomationExtension', False)

    try:
        doc_file = open('C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user\\Session_info.txt', 'r+')
        string_from_file=doc_file.readline()
        if not string_from_file:
            driver = webdriver.Chrome(options=options)
            doc_file.close()
            return driver
        else:
            doc_file.close()
            raise ZeroDivisionError
        # doc_file.write(driver.session_id)


    except:
        time.sleep(30)



def driver_quit(driver):
    try:
        pathname = 'C:\\Users\\yvsobol3\\PycharmProjects\\Test'
        pickle.dump(driver.get_cookies(), open(pathname + "/cookies.pkl", "wb"))
        driver.quit()
        doc_file = open('C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user\\Session_info.txt', 'w+')
        doc_file.close()

    except:
        pass



def get_new_account(Credentials):
    print("Ошибка аунтефикации!")
    print('Введите логин:')
    login = input("")
    print('Введите пароль:')
    password = input("")
    Credentials['login'], Credentials['pass'] = login, password
    save_obj(Credentials, 'Cred')

def add_cookies(driver,path):

    # пробуем взять печеньки из дамп файла
    try:
        cookies = pickle.load(open(path + "/cookies.pkl", "rb"))
        cookie_text_file=open('C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user\\Session_info2.txt', 'w+')
        driver.delete_all_cookies()

        for cookie in cookies:
            cookie_text_file.write(str(cookie)+"\n")
            # if 'expiry' in cookie:
            #     if isinstance(cookie['expiry'], float):
            #         print("старый куки" + cookie['expiry'])
            #         cookie['expiry'] = int(cookie['expiry'])
            #         print("новый куки" + cookie['expiry'])
            driver.add_cookie(cookie)
        cookie_text_file.close()
    except Exception as ex:
        return ex,'нет печенек'
    return 'ok', 'есть печеньки'


def check_login(driver):
    driver.get(
        'http://remedy.msk.mts.ru/arsys/forms/remedy-prom/I2%3AWorks/Default+Administrator+View/?&mode=CREATE')  #
    try:
        text = WebDriverWait(driver, 1).until(EC.visibility_of_element_located(
            (By.ID, "arid_WIN_0_1"))).get_attribute('value')
        if text == "MSK":
            return 'ok', 'получилось'
        else:
            check_login(driver)
    except Exception as ex:
        return ex, 'не получилось'

def login_with_password(driver):
    # cookie=driver.get_cookie('JSESSIONID')
    # driver.delete_all_cookies()
    # driver.add_cookie(cookie)
    driver.get(
        'http://remedy.msk.mts.ru/')
    if os.path.exists("C:\\Users\\yvsobol3\\PycharmProjects\\Test\\cookies.pkl"):
        os.remove("C:\\Users\\yvsobol3\\PycharmProjects\\Test\\cookies.pkl")
        # driver.delete_all_cookies()
        # remedyLogin(driver)
        #
        # get_driver(webdriver.ChromeOptions())
        # driver.delete_all_cookies()
        # driver.get(
        #      'http://remedy.msk.mts.ru/arsys/forms/remedy-prom/I2%3AWorks/Default+Administrator+View/?cacheid=&mode=CREATE')
        # raise NotImplemented

    try:
        WebDriverWait(driver, 1).until(EC.visibility_of_element_located(
            (By.ID, "username-id")))
        try:
            Credentials = load_obj('Cred')
        except:
            print('Введите логин:')
            login = input("")
            print('Введите пароль:')
            password = input("")
            Credentials = {'login': login, 'pass': password}
            save_obj(Credentials, 'Cred')
            remedyLogin(driver)

        login, password = Credentials['login'], Credentials['pass']
        driver.find_element_by_id('username-id').send_keys(login)
        driver.find_element_by_id('pwd-id').send_keys(password)
        driver.find_element_by_id("login").click()
        # теперь здесь должна быть отработка сообщения, о том, что пользователь подключен с другой машины

        # отработка сообщения о том, что пользователь подключен с другой машины, перезаписать?
        try:
            # здесь обрабатывается запрос на перезапись
            alert = driver.switch_to_alert()
            if '' in alert.text:
                alert.accept()
            driver.switch_to_default_content()
            print("Просил перезаписать")
        except:
            ex = 0

        try:
            # здесь рассматривается случай появления сообщения о том, что пользователь уже подключен с другого компьютера
            elem = driver.switch_to.active_element
            driver.switch_to.frame(elem)
            iframe = driver.find_element_by_xpath('//*[@id="PopupMsgBox"]').text
            if iframe.find('9084')!=-1:
                return "Работы не заводятся повторите попытку через 15 минут"
            if iframe.find('9201')!=-1:
                return "Работы не заводятся неверная сессия"
            driver.switch_to_default_content()
        except:
            ex=0


            # print("Не просил перезаписать")
        href = driver.current_url
        href_split = href.split('/')[-1]
        cacheid = href_split.split('=')[-1]
    except Exception as ex:
        return ex,'пароль не сработал'
    return 'ok', 'пароль сработал'

# @retry(tries=3, delay=1.5, backoff=10)
def remedyLogin(driver):
    pathname = 'C:\\Users\\yvsobol3\\PycharmProjects\\Test'
    # try:
    #     cacheid = pickle.load(open(pathname + "/cacheid.pkl", "rb"))
    # except:
    #     cacheid = ''
    n=0
    break_time=10
    while n<5:
        n+=1
        try:
            if check_login(driver)[1] == 'не получилось':

                # if check_login(driver)[1]=='не получилось':
                if add_cookies(driver,pathname)[1]!='нет печенек':
                    if check_login(driver)[1] == 'получилось':
                        # pickle.dump(driver.get_cookies(), open(pathname + "/cookies.pkl", "wb"))
                        return 'получилось'


                if login_with_password(driver)=="Работы не заводятся повторите попытку через 15 минут":
                    raise NotImplemented
                if check_login(driver)[1] == 'не получилось':
                    raise NotImplemented
                else:
                    pickle.dump(driver.get_cookies(), open(pathname + "/cookies.pkl", "wb"))
                    return 'получилось'
            else:
                # pickle.dump(driver.get_cookies(), open(pathname + "/cookies.pkl", "wb"))
                return 'получилось'

        except:
            time.sleep(break_time)
            driver.delete_all_cookies()
            break_time+=60
            pass








        #     try:
        #         # здесь рассматривается случай появления сообщения о том, что пользователь уже подключен с другого компьютера
        #         elem = driver.switch_to.active_element
        #         driver.switch_to.frame(elem)
        #         iframe = driver.find_element_by_xpath('//*[@id="PopupMsgBox"]').text
        #         if iframe.find('9084')!=-1:
        #             return "Работы не заводятся повторите попытку через 15 минут"
        #     except:
        #         ex=0
        #
        #     if 'ARERR [9388]' in driver.find_element_by_xpath('//*[@id="LoginMsg-id"]').text:
        #         get_new_account(Credentials)
        #         remedyLogin(driver)
        #
        #
        # try:
        #     # driver.get(
        #     #     'http://remedy.msk.mts.ru/arsys/forms/remedy-prom/I2%3AWorks/Default+Administrator+View/?cacheid=' + cacheid + '&mode=CREATE')
        #     # попробовать пойти по полученной ссылке с новым кэщ айди
        #     # если все работает сохраняем кэш айди
        #     WebDriverWait(driver, 1).until(EC.visibility_of_element_located(
        #         (By.ID, "arid_WIN_0_536870933")))
        #     href = driver.current_url
        #     href_split = href.split('/')[-1]
        #     href_split = href.split('=')[1]
        #     cacheid = href_split.split('&')[0]
        #     pickle.dump(cacheid, open(pathname + "/cacheid.pkl", "wb"))
        #     pickle.dump(driver.get_cookies(), open(pathname + "/cookies.pkl", "wb"))
        # except Exception as ex:
        #     # удалить файл с куками
        #     if os.path.exists(pathname + "/cookies.pkl"):
        #         os.remove(pathname + "/cookies.pkl")
        #     remedyLogin(driver)
        #     # '/html/body/table[2]/tbody/tr[1]/td[2]/div[2]' ARERR [9506]
        #     print("Сбой авторизации")

def get_performer_by_mail(mail,type_W):
    dicMail = {'stan.oksi@ya.ru': 'Власов Станислав Игоревич',
               'oglm2012@gmail.com': 'Наревский Кирилл Вячеславович',
               'drobotenko_s.a@mail.ru': 'Дроботенко Сергей Алексеевич',
               'aayemely@gmail.com': 'Емельянчиков Андрей Анатольевич',
               'lika_tyutina@bk.ru': 'Тютина Анжелика Александровна',
               # 'vovikst30@gmail.com': 'Стародубцев Владимир Геннадьевич',
               'sob.evg@yandex.ru': "Соболенко Евгений Викторович",
               'mon23@mail.ru': 'Максименко Олег Николаевич',
               'minilink2e@yandex.ru': 'Кузьмин Александр Алексеевич',
               'kitae@mail.ru': 'Смирнов Виктор Викторович',
               'ads@quantumpls.com': 'Шестаков Артем Дмитриевич',
               'cumspe@mail.ru': 'Павлов Дмитрий Витальевич',
               'alex.nastenko0107@gmail.com':'Настенко Александр Петрович',
               'vlad.sokol.95@mail.ru': 'Сокол Владислав Сергеевич',
               'ivashechkinpavel@mail.ru': 'Ивашечкин Павел Владимирович',
               'kokorindmitry.kr@gmail.com':'Кокорин Дмитрий Вадимович',
               'sszavalk@gmail.com': 'Завальковский Сергей Сергеевич',
               'bogachevev@mail.ru':'Богачев Евгений Сергеевич',
               'igor.malkin@solarphon.com': 'Малкин Игорь Юльевич',
               # 'ivk0750708@gmail.com': 'Иван Иванович Комисаренко',
                'dimat2006@mail.ru': 'Трепалин Дмитрий Александрович'}
    dicGroup = {'mon23@mail.ru': 'ЮГ\Краснодар\ОРС\ТС',
                'vovikst30@gmail.com': 'ЮГ\Краснодар\Подрядные организации\ООО Инсол Телеком',
                'sob.evg@yandex.ru': 'ЮГ\Краснодар\ОРС\ТС'}
    dicIniciator = {'minilink2e@yandex.ru': 'Кузьмин Александр Алексеевич',
                    'alex.nastenko0107@gmail.com': 'Настенко Александр Петрович',
                    'oglm2012@gmail.com': 'Наревский Кирилл Вячеславович',
                    'sszavalk@gmail.com': 'Завальковский Сергей Сергеевич',
                    'ivashechkinpavel@mail.ru': 'Ивашечкин Павел Владимирович',
                    'mon23@mail.ru': 'Кузьмин Александр Алексеевич'
                    }
    dicIniciatorBS = {'stan.oksi@ya.ru': 'Сытник Алексей Викторович',
                    'dimat2006@mail.ru': 'Сытник Алексей Викторович',
                    'vlad.sokol.95@mail.ru': 'Сытник Алексей Викторович',
                    'igor.malkin@solarphon.com': 'Сытник Алексей Викторович',
                      'drobotenko_s.a@mail.ru':  'Сытник Алексей Викторович'
                      }
    try:
        performer = dicMail[mail]
    except:
        performer='error'
        print('Пользователя ' + mail + ' нет в базе')



    try:
        performerG = dicGroup[mail]
    except:
        performerG = ''
    if type_W=='RN':
        try:
            iniciator = dicIniciatorBS[mail]
        except:
            iniciator = ''
    else:
        try:
            iniciator = dicIniciator[mail]
        except:
            iniciator = ''
    return performer,performerG, iniciator

def getParam(mail_body):
    List_strings=[]
    List_strings_proto =mail_body.split('\n')
    for string in List_strings_proto:
        Sub_string_list=string.split('\r')
        for sub_string in Sub_string_list:
            if len(sub_string)>7: #рассчитываем, что минимум должно быть 5 символов в номере БС, пробел и один символ описания
                BS=string.split(' ')[0].strip(' _-').replace('_','').replace('-','')
                BSR = ""
                for n in BS[::-1]:
                    if n.isdigit():
                        BSR = n + BSR
                if BS==BSR:
                    List_strings.append(sub_string.strip(' _-'))

    List_BS_with_describe=[]
    for string in List_strings:
        try:
            BS=string.split(' ')[0].strip(' _-').replace('_','').replace('-','')
            describeW=""
            for str1 in string.split(' ')[1:]:
                describeW=describeW+str1+" "
            BSR=""
            for n in BS[::-1]:
                if n.isdigit():
                    BSR = n + BSR
            BSR = BSR[:2] + "_" + BSR[2:]
            describeW=describeW.split('\n')[0]
            describeW = describeW.split('\r')[0]
        except:
            print("Не могу распознать строку")
            return 'error'
        List_BS_with_describe.append([BSR,describeW])
    return List_BS_with_describe



def sel_NE(NE_list,BS,typeNE):
    if typeNE=='БС':
        try:
            for n, we in enumerate(NE_list):
                if BS + '_G' in we.text:
                   return NE_list[n]
            for n, we in enumerate(NE_list):
                if BS + '_D' in we.text:
                   return NE_list[n]
            for n, we in enumerate(NE_list):
                if BS + '_GD' in we.text:
                   return NE_list[n]
            for n, we in enumerate(NE_list):
                if BS + '_U' in we.text:
                    return NE_list[n]
            for n, we in enumerate(NE_list):
                if 'BTS_'+BS in we.text:
                   return NE_list[n]
            # for n, we in enumerate(NE_list):
            #     if 'MBH_'+BS in we.text:
            #        return NE_list[n]
            for n, we in enumerate(NE_list):
                if "PL_"+BS in we.text:
                    return NE_list[n]
            return NE_list[0]
        except:
            print("title", "Не получилось найти БС")
            return 'error'
    elif typeNE=='РРС':
        try:
            for n, we in enumerate(NE_list):
                if BS + '_3' in we.text:
                   return NE_list[n]
            for n, we in enumerate(NE_list):
                if BS + '_2' in we.text:
                   return NE_list[n]
            for n, we in enumerate(NE_list):
                if BS + '_1' in we.text:
                   return NE_list[n]
            for n, we in enumerate(NE_list):
                if BS  in we.text:
                    return NE_list[n]
            return NE_list[0]
        except:
            print("title", "Не получилось найти БС")
            return 'error'
    elif typeNE=='MBH':
        try:
            # for n, we in enumerate(NE_list):
            #     if 'MBH_'+BS in we.text:
            #        return NE_list[n]
            for n, we in enumerate(NE_list):
                if BS + '_G' in we.text:
                   return NE_list[n]
            for n, we in enumerate(NE_list):
                if BS + '_D' in we.text:
                   return NE_list[n]
            for n, we in enumerate(NE_list):
                if BS + '_U' in we.text:
                    return NE_list[n]
            for n, we in enumerate(NE_list):
                if "PL_"+BS  in we.text:
                    return NE_list[n]
            for n, we in enumerate(NE_list):
                if BS in we.text:
                   return NE_list[n]
            return NE_list[0]
        except:
            print("title","Не получилось найти БС",BS)
            return 'error'
def sel_classificator(NE_list,value):
    try:
        for n, we in enumerate(NE_list):
            if value in we.text:
               return NE_list[n]
    except:
        messagebox.showinfo("title","Не получилось найти БС")
        return 'error'

def add_NE_impact(driver,NE):
    for ne in NE:
        if not '_' in ne:
            ne=ne[:2]+'_'+ne[2:]
        set_value(driver, "arid_WIN_0_536888676", '%'+ne+'%')
        exist_name_list=[]
        nioss_el = driver.find_elements_by_xpath('//*[@id="T778000008"]/tbody/tr/td[2]/nobr/span')
        for el in nioss_el:
            exist_name_list.append(el.get_attribute('innerText'))
    #     выбираем все значения из таблицы с сетевыми элементами
        nioss_el = driver.find_elements_by_xpath('//*[@id="T536870926"]/tbody/tr/td[2]/nobr/span')
        for el in nioss_el:
            # el=el.find_element_by_xpath('//td[2]/nobr/span')
            if 'BTS' in el.get_attribute('innerText') and not(el.get_attribute('innerText') in exist_name_list):
                el.click()
                driver.find_element_by_xpath('//*[@id="WIN_0_536870891"]/div/div').click()

# @retry(EX.NoSuchElementException, tries=15,delay=5,backoff=2)
def get_work_number(driver):
    time.sleep(5)

    try:
        driver.switch_to_default_content()
        driver.switch_to.window(driver.window_handles[0])
        el=driver.find_elements_by_id("arid_WIN_0_1")
        print('количество найденных элементов:'+str(len(el)))
        work_number = el[0].get_attribute('value')
        if work_number=="" or work_number=='MSK':
            time.sleep(4)
            el = driver.find_elements_by_id("arid_WIN_0_1")
            work_number = el[0].get_attribute('value')
        print("Номер работ:",work_number)
    except:
        try:
            time.sleep(14)
            try:
                elem = driver.switch_to.active_element
                driver.switch_to.frame(elem)
                iframe = driver.find_element_by_xpath('//*[@id="PopupMsgBox"]').text

                print("!!!"+iframe+'!!!')
                alert = driver.switch_to_alert()
                alert.accept()
                driver.switch_to_default_content()
            except:
                pass
            driver.switch_to.window(driver.window_handles[0])

            el = driver.find_elements_by_id("arid_WIN_0_1")
            work_number = el[0].get_attribute('value')
        except:
            print("Ожидание не помогло, номер работы не считан")
            work_number="Не считано!"

    return work_number

@retry(tries=3, delay=13)
def set_klassificator(driver,klassificator):
    driver.find_element_by_xpath('//*[@id="reg_img_536870872"]').click()
    driver.switch_to.window(driver.window_handles[1])

    for n, sub_klass in enumerate(klassificator):
        we_table = driver.find_elements_by_xpath('//*[@id="T536871022"]/tbody/tr')
        ne = sel_classificator(we_table, '.'.join(klassificator[:n + 1]))
        actionChains = ActionChains(driver)
        actionChains.double_click(ne).perform()
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(1)
    # if driver.window_handles[1]:
    #     raise NotImplemented

def set_value_initiator(driver,initiator):
    if initiator:
        if ('Максименко' in
            initiator):  # or ( 'Кузьмин' in kwargs['initiator']):  # защита от однофамильцев Максименко Олега и Кузьмина Саши
            set_value(driver, "arid_WIN_0_540000008",initiator)
            driver.switch_to.window(driver.window_handles[1])
            elements = driver.find_elements_by_xpath('//*[@id="T536870926"]/tbody/tr/td[2]')
            for el in elements:
                if 'ведущий' in el.text:
                    actionChains = ActionChains(driver)
                    actionChains.double_click(el).perform()
                    driver.switch_to.window(driver.window_handles[0])
                    break
        else:
            set_value(driver, "arid_WIN_0_540000008", initiator)

def remedy_set_param(driver,**kwargs):
    # n=1
    # while n<4:
    #     n+=1
    #     try:
    if len(kwargs['BS'].split("<>"))>1: kwargs['BS']=kwargs['BS'].split("<>")[0] #отрабатывает случай, если вместо БС введена РРЛ вида А<>B

    # if ('Максименко' in kwargs['performer']) : #защита от однофамильцев Максименко Олега и Кузьмина Сашиor ('Кузьмин' in kwargs['performer'])
    #     set_value(driver, "arid_WIN_0_536870933", kwargs['performer'])
    #     driver.switch_to.window(driver.window_handles[1])
    #     elements=driver.find_elements_by_xpath('//*[@id="T536870926"]/tbody/tr/td[2]')
    #     for el in elements:
    #         if 'ведущий' in el.text:
    #             actionChains = ActionChains(driver)
    #             actionChains.double_click(el).perform()
    #             driver.switch_to.window(driver.window_handles[0])
    #             break
    # elif 'Смирнов' in kwargs['performer']:  # защита от однофамильцев Витя Смирнов
    #         set_value(driver, "arid_WIN_0_536870933", kwargs['performer'])
    #         driver.switch_to.window(driver.window_handles[1])
    #         elements = driver.find_elements_by_xpath('//*[@id="T536870926"]/tbody/tr/td[2]')
    #         for el in elements:
    #             if 'инженер' in el.text:
    #                 actionChains = ActionChains(driver)
    #                 actionChains.double_click(el).perform()
    #                 driver.switch_to.window(driver.window_handles[0])
    #                 break
    # else:
    #общая информация
    set_value(driver, "arid_WIN_0_536870919", kwargs['performerG']) #группа исполнителя
    set_value(driver, "arid_WIN_0_536870933", kwargs['performer']) #исполнитель
    set_value(driver, "arid_WIN_0_536870923", kwargs['supervisor']) #контролирующий
    set_value_initiator(driver,kwargs['initiator']) #инициатор
    if kwargs['incident']: set_value(driver, "arid_WIN_0_540000017", kwargs['incident']) #если есть инцидент, то заполняем и его

    # даты
    startW = (datetime.datetime.now() + datetime.timedelta(minutes=20)).strftime('%d.%m.%Y %H:%M:%S')
    endW = (datetime.datetime.now() + datetime.timedelta(days=kwargs['days'])).strftime('%d.%m.%Y %H:%M:%S')
    set_value(driver, "arid_WIN_0_820365", startW)
    set_value(driver, "arid_WIN_0_820366", endW)
    # set_value(driver, "arid_WIN_0_540000045", startW)
    # set_value(driver, "arid_WIN_0_537000005", endW)


    # заполняем номер БС
    driver.switch_to.window(driver.window_handles[0])
    driver.find_element_by_xpath('//*[@id="reg_img_536870900"]').click()
    driver.switch_to.window(driver.window_handles[1])
    set_value(driver, "arid_WIN_0_536870935", kwargs['BS'])
    we_table = driver.find_elements_by_xpath('//*[@id="T536870924"]/tbody/tr')

    actionChains = ActionChains(driver)
    ne=sel_NE(we_table, kwargs['BS'],kwargs['typeNE'])
    if ne!='error':
        actionChains.double_click(ne).perform()
    else:
        return 'Ошибка! Не удалось найти СЭ в Remedy.'

    driver.switch_to.window(driver.window_handles[0])

    # влияние на сервис начало/конец
    if kwargs['impactS'] !='Нет влияния на сервис абонентам в зоне действия СЭ.':
        endW = (datetime.datetime.now() + datetime.timedelta(days=1)).strftime('%d.%m.%Y %H:%M:%S')
        set_value(driver, "arid_WIN_0_540000045", startW)
        set_value(driver, "arid_WIN_0_537000005", endW)
        set_value(driver, "arid_WIN_0_536870939", startW)
        set_value(driver, "arid_WIN_0_536870942", endW)

    # try:
    time.sleep(0.5)
    set_value(driver, "arid_WIN_0_536870930", kwargs['typeW'])
    time.sleep(0.5)
    set_value(driver, "arid_WIN_0_8", kwargs['describeW'])
    set_value(driver, "arid_WIN_0_536870956", kwargs['typeS'])
    set_value(driver, "arid_WIN_0_540000012", kwargs['impactS'])
    set_value(driver, "arid_WIN_0_537000124", kwargs['addInfo'])





    # заполняем классификатор
    set_klassificator(driver,kwargs['klassificator'])


    if kwargs['opercontext']: set_value(driver, "arid_WIN_0_540000049", kwargs['opercontext'])

    # жмем кнопку сохранить
    driver.find_element_by_xpath('//*[@id="WIN_0_536870907"]/div/div').click()
    time.sleep(3)
    # если появилось окно о информировании жмем ОК
    work_number=get_work_number(driver)
    if work_number == 'MSK':
        time.sleep(10)
        work_number=work_number+'+'+get_work_number(driver)
    try:
        if kwargs['impactS'] != 'Нет влияния на сервис абонентам в зоне действия СЭ.':
            time.sleep(2.0)
            try:
                alert = driver.switch_to_alert()
                alert.accept()
                driver.switch_to_default_content()
            except:
                driver.refresh()
            time.sleep(5.5)
            # добавляем ресурвы в свисок с влиянием
            driver.find_element_by_xpath(
                '// *[@id = "WIN_0_1000000050"]/div[2]/div[2]/div/dl/dd[6]/span[2]/a').click()
            driver.find_element_by_xpath('//*[@id="WIN_0_1000004000"]/div/div').click()
            time.sleep(1.5)
            elem = driver.switch_to.active_element
            # iframe = driver.find_element_by_xpath("//iframe[@id='1613737641694P']")
            driver.switch_to.frame(elem)
            driver.find_element_by_xpath('//*[@id="arid_WIN_0_1000000060"]').clear()
            add_NE_impact(driver, kwargs['NE_impact'])
            time.sleep(1.0)
            driver.find_element_by_xpath('//*[@id="WIN_0_778000011"]/div/div').click()
            time.sleep(2)
            driver.find_element_by_xpath('//*[@id="WIN_0_536870907"]/div/div').click()

    except Exception as ex:
        print('Ошибка %s в блоке привязки ресурсов к работе с перерывом!' %ex)
    # time.sleep(60)

    return work_number
        # except:
        #     remedyLogin(driver)

    # try:
    #     driver.switch_to_frame(0)
    #     el=driver.find_element_by_xpath(".//*")
    #     el.send_keys(Keys.RETURN)
    # except:
    #     print('Работа заведена без уведомлений')
def reply_send(to_address,to_name,subject,text):
    # инициализируем объект outlook
    olk = win32com.client.Dispatch("Outlook.Application")
    Msg = olk.CreateItem(0)

    # формируем письма, выставляя адресата, тему и текст
    Msg.To = to_address
    Msg.Subject =  subject  # добавляем RE в тему

    Msg.Body = "Здравствуйте, " + str(to_name) + "! \n\n" \
                                                 "Ваш запрос отработан! \n"+text+\
                                                 "\n\n\n\n\n\n****справка по использованию автозаведения****" + \
                                                 "\n ********************************************" + \
               "\nping - в теме письма укажите ping в сообщении ip, в ответе будет информация о доступности " + \
               "\nданные по площадке - в теме письма укажите getidbs в сообщении номер БС" + \
               "\nданные по пролету + ТЗ - в теме письма укажите getidrrl в сообщении пролет" + \
               "\nданные по чеклисту - в теме письма укажите checklistrrl в сообщении пролет" + \
               "\nДля автозаведения работ необходимо отправить сообщение с " +\
                                                 "\nавторизованного адреса на yvsobol3 с темой newwork " +\
                                                 "\nс текстом: RRNNNN описание работ, где RR-номер региона, NNNN-номер БС." + \
                                                  "\nЕсли в тексте есть слова осмотр, юстировка или демонтаж, то выбирается соответствующий классификатор," + \
                                                   "\nв других случаях выбирается классификатор :монтаж оборудования." + \
               "\nДля автозаведения работ на УИТС необходимо отправить сообщение с " + \
               "с темой newwork УИТС" + \
               "\nс текстом: RRNNNN описание работ, где RR-номер региона, NNNN-номер БС." + \
               "\nЕсли в тексте есть слова переключение, выделение или ВОЛС, то выбирается соответствующий классификатор," + \
               "\nна переключение нагрузки, выделение адресации или включение сегмента ВОЛС." +\
               "\nв других случаях выбирается классификатор: Оптимизация параметров и настроек ТС." + \
               "\nДля автозаведения работ с перерывом сервиса необходимо отправить сообщение с " + \
               "\nс темой, содержащей !" + \
               "\nЕсли кроме указанной в письме БС, будут выключаться другие: необходимо их перечислить после!" + \
               "\nв формате !RRNNNN RRNNNN RRNNNN" + \
               "\nПереключение оформляется на суточный срок после заведения работ." + \
               "\nВ случае, если во влияние нужно добавить только текущую БС, перечислять БС после ! не нужно."+\
                 "\nЧтобы заводить сразу несколько работ в одном письме, информацию о новой работе вводите отдельной строкой."
        # и отправляем
    Msg.Send()


def work_create(driver,arg1,arg2,arg3):
    # sys.argv = ['C:\\Users\\yvsobol3\\PycharmProjects\\Test\\Remedy_web_1_0.py', '232596 переключение нагрузки ** days=5\r \n', 'sob.evg@yandex.ru', 'newwork !']

    try:
        if ('bs' in arg3.lower()) or ('бс' in arg3.lower()):
            performer, performerG, initiator = get_performer_by_mail(arg2,'RN')
        else:
            performer, performerG, initiator = get_performer_by_mail(arg2,'TN')
        mail_person=performer
        List_BS_with_describe =getParam(arg1)
        if performerG=='error':
            reply_send(arg2, 'Пользователь', 'Автозаведение работ', 'Email не зарегистрирован!')
            return 'error','Email не зарегистрирован!'
        if not List_BS_with_describe:
            reply_send(arg2, 'Пользователь', 'Автозаведение работ', 'Не удалось распознать сообщение.')
            return 'error', 'Не удалось распознать сообщение.'
    except:
        print('Ошибка распознавания в письме')
        reply_send(arg2, 'Пользователь', 'Автозаведение работ', 'Ошибка распознавания текста в письме!')
        return 'error', 'Ошибка распознавания текста в письме!'

    for BS,describeW in List_BS_with_describe:
        NE_impact = []
        start_time = time.time()




        # time.sleep(1000)
        typeW = 'Планово-профилактическая'
        incident=''
        impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
        klassificator = ['02', '04']
        addInfo = ''
        opercontext=''
        supervisor = 'Оперативный дежурный ООКУ СРД ФО'
        days = 1
        typeNE = 'БС'

        if '**' in describeW:
            days_count=''
            describeW,serv_str=describeW.split('**',1)
            if "days" in serv_str:
                for n in serv_str[::-1]:
                    if n.isdigit():
                        days_count = n + days_count
            if int(days_count)!="0":
                days=int(days_count)



        # addInfo = "На основании п 1.3 Распоряжения №02/00172р от 02.09.2021 «О введении моратория на период проведения выборов депутатов Государственной Думы». "
        # в этой секции анализируется текст сообщения для того, чтобы по словам маркерам выбрать классификатор
        # по умолчанию классификатор "монтаж оборудования"
        if 'осмотр' in describeW.lower():
            klassificator = ['04', '12']

        if 'демонтаж' in describeW.lower():
            klassificator = ['02', '05']

        if 'юстиров' in describeW.lower():
            klassificator = ['02', '19']

        if 'измерени' in describeW.lower():
            klassificator = ['03', '07']

        if 'агос' in describeW.lower():
            klassificator = ['15', '04', '07']
            performerG = "ЮГ\Краснодар\ЭТС\TN"
            performer = ""
        # конец секции



        # если в теме письма содержится УИТС, то работа заводится на УИТС
        # отрабатываются только конкретные сочетания слов в описании работы

        if 'уитс' in arg3.lower():
            days = 3
            klassificator = ['14', '05', '07']
            if 'переключени' in describeW.lower():
                impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
                klassificator = ['14', '01','10']

                addInfo = addInfo + " Локально работы по переключению выполняет " + performer + '.'
                addInfo = addInfo + 'Прошу сообщить регистратору список СЭ для включения в список объектов с влиянием.'
                days = 2
                typeNE = 'БС'
            elif 'выделени' in describeW.lower():
                impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
                klassificator = ['14', '05', '02']
                days = 2
                typeNE = 'БС'
                addInfo = addInfo + " Работу по выделению ip адреса инициировал " + performer + '.'
            elif 'волс' in describeW.lower():
                impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
                klassificator = ['14', '01', '09']
                days = 2
                typeNE = 'БС'
                addInfo = addInfo + " Работы по сборке трассы ВОЛС выполняет " + performer + '.'
            elif 'чек' in describeW.lower() and 'лист' in describeW.lower():
                impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
                klassificator = ['14', '01', '06']
                days = 3
                typeNE = 'БС'
                addInfo = addInfo + " Работу по проверке чек-листа инициировал " + performer + '.'
            elif 'лицензи' in describeW.lower():
                impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
                klassificator = ['14', '02', '02']
                days = 3
                typeNE = 'БС'
                addInfo = addInfo + " Работу по заливке лицензий инициировал " + performer + '.'
            else:
                addInfo = addInfo + " Работу инициировал " + performer + '.'

            performerG = 'ЮГ\ЕЦУС\УИТС\MBH Юго-Запад'
            performer = ''

        if 'омод' in arg3.lower():

            klassificator = ['01', '03', '03','08']
            if 'tcu' or 'siu' in describeW.lower():
                impactS = 'Незначительное/кратковременное влияние на сервис абонентам в зоне действия СЭ.'
                klassificator = ['01', '03', '03','08']
                addInfo = addInfo + " Работы по монтажу выполняет " + performer + '.'
                days = 2
                typeNE = 'БС'
            performerG = 'ЮГ\ЕЦУС\ОМОД\Гр КиР СРД'
            performer = ''
            opercontext='Ericsson 2G'

        # по просьбе эксплуатации, если в теме письма содержится ДЭ
        # то заводится работа на классификатор эксплуатации

        if 'дэ' in arg3.lower():
            klassificator = ['15', '12', '01']

        # по просьбе эксплуатации, если в теме письма содержится ДЭ
        # то заводится работа на классификатор эксплуатации
        if 'avr' in arg3.lower():
            klassificator = ['15', '03', '02']
            typeW = 'Аварийно-восстановительная'
            words_list=describeW.lower().split(" ")
            for word in words_list:
                if 'msk' in word:
                    incident = word
        # отрабатываем случай если работа заводится с влиянием на сервис.
        # формат темы сообщения должен быть *newwork!231252 230025*

        if '!' in arg3:
            impactS = 'Незначительное/кратковременное влияние на сервис абонентам в зоне действия СЭ.'
            NE_impact.append(BS)
            BS_impact_list= arg3.split('!')[-1]
            BS_impact_list =str(BS_impact_list).strip()
            BS_impact_list = BS_impact_list.split(' ')
            for BS_impact in BS_impact_list:
                if BS_impact:
                    NE_impact.append(BS_impact)




        text=("Работа с входными данными: номер БС %s" %BS)
        text = text +("\nТип СЭ: %s, Описание: %s, Влияние: %s" % (typeNE,describeW,impactS))
        text = text +("\nИсполнитель: %s, Группа исполнителя: %s" % (performer, performerG))
        if NE_impact:
            text = text + ("\nЗатронутые БС в работе:" +', '.join(NE_impact))
            if (days > 1) :
                days=1
                text = text + ("\nРабота с влиянием заводится только на 1 сутки. Срок скорректирован.")
        print(text)
        text=text + "\nЗаведена работа № %s" %remedy_set_param(driver,typeW=typeW,
        typeS='Данные + Голос',
        impactS=impactS,
        addInfo=addInfo,
        supervisor=supervisor,
        incident=incident,
        initiator=initiator,
        BS=BS, describeW=describeW,performer=performer, performerG=performerG,
        typeNE=typeNE,klassificator=klassificator,
                         days=days,NE_impact=NE_impact,opercontext=opercontext)
        text=text+"\n"


        print(text+"\nВремя выполнения программы: %s секунд" % (time.time() - start_time))
        reply_send(arg2,mail_person,'Автозаведение работ',text)
        check_login(driver)



