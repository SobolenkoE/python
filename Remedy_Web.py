from selenium import webdriver
import pickle
import sys
from NIOSS import set_value
from tkinter import messagebox
# import Main
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from retry import retry
from selenium.common.exceptions import WebDriverException
import win32com.client
import re, os, time, datetime, uuid

Credentials = {'login': 'yvsobol3', 'pass': 'Cj,jk115'}

def remedyLogin(driver,login,password):
    pathname = os.path.dirname(sys.argv[0])
    driver.delete_all_cookies()
    driver.get('http://remedy.msk.mts.ru/') # для того, чтобы появилась возможность прикрутить куки
    cookies = pickle.load(open(pathname+"/cookies.pkl", "rb"))
    for cookie in cookies:
        driver.add_cookie(cookie)
    driver.get('http://remedy.msk.mts.ru/arsys/forms/remedy-prom/I2%3AWorks/Default+Administrator+View/?cacheid=&mode=CREATE')
    try:
        driver.find_element_by_id('username-id').send_keys(login)
        driver.find_element_by_id('pwd-id').send_keys(password)
        driver.find_element_by_id("login").click()
        try:
            alert = driver.switch_to_alert()
            print(alert)
            alert.accept()
            driver.switch_to_default_content()
        except:
            print("залогинились после ввода логина и пароля")
        pickle.dump(driver.get_cookies(), open(pathname+"/cookies.pkl", "wb"))
    except:
        print("Логиниться не нужно(зашли по кукам)")
        try:
            alert = driver.switch_to_alert()
            alert.accept()
            driver.switch_to_default_content()
        except:
            print("залогинились по куки")

def getParam(args):
    dicMail={'stan.oksi@ya.ru':'Власов Станислав Игоревич','vovikst30@gmail.com':'Стародубцев Владимир Геннадьевич',
             'sob.evg@yandex.ru':"Соболенко Евгений Викторович",'mon23@mail.ru':'Максименко Олег Николаевич','minilink2e@yandex.ru':'Кузьмин Александр Алексеевич','kitae@mail.ru':'Смирнов Виктор Викторович','ivk0750708@gmail.com':'Иван Иванович Комисаренко'}
    dicGroup = {'mon23@mail.ru': 'ЮГ\Краснодар\ОРС\ТС','vovikst30@gmail.com':'ЮГ\Краснодар\Подрядные организации\ООО Инсол Телеком','sob.evg@yandex.ru': 'ЮГ\Краснодар\ОРС\ТС'}
    try:
        performer=dicMail[args[2]]
    except:
        messagebox.showinfo('Title','Пользователя ' + args[2] +' нет в базе')
        sys.exit()

    try:
        performerG = dicGroup[args[2]]
    except:
        performerG =''

    try:
        BS=args[1].split(' ')[0]
        describeW=""
        for str1 in args[1].split(' ')[1:]:
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

    return BSR,describeW,performer,performerG


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
                if "PL_"+BS  in we.text:
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
            for n, we in enumerate(NE_list):
                if 'MBH_'+BS in we.text:
                   return NE_list[n]
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
            print("title","Не получилось найти БС")
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



def remedy_set_param(driver,**kwargs):
    # возможно появляется окно о том, что пользовател подключен с другого компьютера, проверяем
    # try:
    # driver.find_element_by_xpath('//*[@id="PopupMsgFooter"]/a').click()
    #
    # time.sleep(60)
    # remedy_set_param(driver, **kwargs)
    # sys.exit()
    # # except:
    # print("с пользователем все хорошо")
    # вкладка назначение
    if len(kwargs['BS'].split("<>"))>1:
        kwargs['BS']=kwargs['BS'].split("<>")[0]
    set_value(driver, "arid_WIN_0_536870919",kwargs['performerG'])
    set_value(driver, "arid_WIN_0_536870933", kwargs['performer'])
    # if kwargs['performer']:
    #     nioss_el = driver.find_element_by_id('arid_WIN_0_536870933')
    #     nioss_el.clear()
    #     nioss_el.send_keys(kwargs['performer'])

    set_value(driver, "arid_WIN_0_536870923", kwargs['supervisor'])
    # set_value(driver,"arid_WIN_0_540000008",initiator)

    # даты
    startW = (datetime.datetime.now() + datetime.timedelta(minutes=20)).strftime('%d.%m.%Y %H:%M:%S')
    endW = (datetime.datetime.now() + datetime.timedelta(days=kwargs['days'])).strftime('%d.%m.%Y %H:%M:%S')
    set_value(driver, "arid_WIN_0_540000045", startW)
    set_value(driver, "arid_WIN_0_537000005", endW)


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

    # try:
    time.sleep(0.5)
    set_value(driver, "arid_WIN_0_536870930", kwargs['typeW'])
    set_value(driver, "arid_WIN_0_8", kwargs['describeW'])
    set_value(driver, "arid_WIN_0_536870956", kwargs['typeS'])
    set_value(driver, "arid_WIN_0_540000012", kwargs['impactS'])
    set_value(driver, "arid_WIN_0_537000124", kwargs['addInfo'])

    # влияние на сервис начало/конец
    if kwargs['impactS'] !='Нет влияния на сервис абонентам в зоне действия СЭ.':
        set_value(driver, "arid_WIN_0_536870939", startW)
        set_value(driver, "arid_WIN_0_536870942", endW)


    # заполняем классификатор
    driver.find_element_by_xpath('//*[@id="reg_img_536870872"]').click()
    driver.switch_to.window(driver.window_handles[1])


    for n,sub_klass in enumerate(kwargs['klassificator']):
        we_table = driver.find_elements_by_xpath('//*[@id="T536871022"]/tbody/tr')
        ne=sel_classificator(we_table,'.'.join(kwargs['klassificator'][:n+1]))
        actionChains = ActionChains(driver)
        actionChains.double_click(ne).perform()

    driver.switch_to.window(driver.window_handles[0])


    # жмем кнопку сохранить
    driver.find_element_by_xpath('//*[@id="WIN_0_536870907"]/div/div').click()

    # если появилось окно о информировании жмем ОК
    try:
        if kwargs['impactS'] != 'Нет влияния на сервис абонентам в зоне действия СЭ.':
            time.sleep(2.5)
            try:
                alert = driver.switch_to_alert()
                alert.accept()
                driver.switch_to_default_content()
            except:
                driver.refresh()
            # добавляем ресурвы в свисок с влиянием
            print(driver.find_element_by_id('arid_WIN_0_1').get_attribute('value'))
            time.sleep(1.5)
            driver.find_element_by_xpath(
                '// *[@id = "WIN_0_1000000050"]/div[2]/div[2]/div/dl/dd[5]/span[2]/a').click()
            driver.find_element_by_xpath('//*[@id="WIN_0_1000004000"]/div/div').click()
            time.sleep(1.5)
            driver.switch_to.window(driver.window_handles[1])
            driver.find_element_by_xpath('//*[@id="arid_WIN_0_1000000060"]').clear()
            add_NE_impact(driver, kwargs['NE_impact'])
            time.sleep(1.0)
            driver.find_element_by_xpath('//*[@id="WIN_0_778000011"]/div/div').click()
    except Exception as ex:
        print('Ошибка %s в блоке привязки ресурсов к работе с перерывом!' %ex)
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(3.0)
    # return driver.find_element_by_id('arid_WIN_0_1').get_attribute('value')
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
                                                 "\n\n\n\n\n\n***справка по использованию автозаведения***" + \
                                                 "\n ********************************************" + \
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
               "\nВ случае, если во влияние нужно добавить только текущую БС, перечислять БС после ! не нужно."

        # и отправляем
    Msg.Send()


if __name__ == "__main__":
    # sys.argv = ['C:\\Users\\yvsobol3\\PycharmProjects\\Test\\Remedy_web.py', '231016 демонтаж ррл\r \n', 'sob.evg@yandex.ru', 'newwork !231016']

    try:
        BS, describeW, performer,performerG =getParam(sys.argv)
    except:
        print('Ошибка распознавания в письме')
        reply_send(sys.argv[2], 'Пользователь', 'Автозаведение работ', 'Ошибка распознавания текста в письме!')
        sys.exit()



    NE_impact = []
    start_time = time.time()
    options = webdriver.ChromeOptions()
    options.add_argument('--profile-directory=Default')
    options.add_argument('headless') #если хотим запустить chrome недивимкой
    options.add_argument("window-size=1920x1080")
    driver = webdriver.Chrome(options=options)

    remedyLogin(driver, Credentials['login'],Credentials['pass'])

    impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
    klassificator = ['02', '04']
    addInfo = ''
    supervisor = 'Оперативный дежурный ООКУ СРД ФО'
    days = 1
    typeNE = 'БС'
    # в этой секции анализируется текст сообщения для того, чтобы по словам маркерам выбрать классификатор
    # по умолчанию классификатор "монтаж оборудования"
    if 'осмотр' in sys.argv[1].lower():
        klassificator = ['04', '12']

    if 'демонтаж' in sys.argv[1].lower():
        klassificator = ['02', '05']

    if 'юстиров' in sys.argv[1].lower():
        klassificator = ['02', '19']

    # конец секции

    # если в теме письма содержится УИТС, то работа заводится на УИТС
    # отрабатываются только конкретные сочетания слов в описании работы

    if 'уитс' in sys.argv[3].lower():
        klassificator = ['14', '05', '06']
        if 'переключени' in describeW.lower():
            impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
            klassificator = ['14', '01','04']
            # addInfo ="На основании п 1.3 Распоряжения о введении моратория. "
            addInfo = addInfo + 'Прошу сообщить регистратору список СЭ для включения в список объектов с влиянием.'
            days = 2
            typeNE = 'БС'
        elif 'выделени' in describeW.lower():
            impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
            klassificator = ['14', '01', '02']
            days = 2
            typeNE = 'БС'
        elif 'волс' in describeW.lower():
            impactS = 'Нет влияния на сервис абонентам в зоне действия СЭ.'
            klassificator = ['14', '02', '01']
            days = 2
            typeNE = 'БС'
        addInfo=addInfo+" Работы по монтажу выполняет " +  performer +'.'
        performerG = 'ЮГ\ЕЦУС\УИТС\MBH Запад'
        performer = ''

    # по просьбе эксплуатации, если в теме письма содержится ДЭ
    # то заводится работа на классификатор эксплуатации

    if 'дэ' in sys.argv[3].lower():
        klassificator = ['15', '12', '01']


    # отрабатываем случай если работа заводится с влиянием на сервис.
    # формат темы сообщения должен быть *newwork!231252 230025*

    if '!' in sys.argv[3]:
        impactS = 'Незначительное/кратковременное влияние на сервис абонентам в зоне действия СЭ.'
        NE_impact.append(BS)
        BS_impact_list= sys.argv[3].split('!')[-1]
        BS_impact_list =str(BS_impact_list).strip()
        BS_impact_list = BS_impact_list.split(' ')
        for BS_impact in BS_impact_list:
            NE_impact.append(BS_impact)

    text=("Работа с входными данными: номер БС %s" %BS)
    text = text +("\nТип СЭ: %s, Описание: %s, Влияние: %s" % (typeNE,describeW,impactS))
    text = text +("\nИсполнитель: %s, Группа исполнителя: %s" % (performer, performerG))
    print(text)
    text="Заведена работа №" +remedy_set_param(driver,typeW='Планово-профилактическая',
    typeS='Данные + Голос',
    impactS=impactS,
    addInfo=addInfo,
    supervisor=supervisor,
    initiator='',
    BS=BS, describeW=describeW,performer=performer, performerG=performerG,
    typeNE=typeNE,klassificator=klassificator,
                     days=days,NE_impact=NE_impact)+"\n"+text


    print(text+"\nВремя выполнения программы: %s секунд" % (time.time() - start_time))
    reply_send(sys.argv[2],performer,'Автозаведение работ',text)
    driver.quit()
    sys.exit()