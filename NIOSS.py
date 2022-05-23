import time
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from retry import retry

class NIOSS_RRL:
    def __init__(self):
        self.Name = ""
        self.PL_A = ""
        self.PL_B = ""
        self.H1_A = ""
        self.H1_B = ""
        self.D1_A = ""
        self.D1_B = ""



# фунция вытаскивающая Object Id из URL
def getObgectID(URL,separator):
    # https://nioss/ncobject.jsp?id=9153410135213621294&site=6031643620013350054
    # https:/nioss/ncobject.jsp?id=9153410135213621294
    ID=''
    URL_list = URL.split('&')
    for URL in URL_list:
        if separator+'=' in URL:
            ID=''
            for n in URL[::-1]:
                if n.isdigit():
                    ID=n+ID
    if len(ID)<10:
        ID=getObgectID(URL, 'id')
    return ID

# фунция удаляющая нули и точки в конце
def StripFreqOB(Freq):
    Freq=str(Freq)
    ID=Freq
    # try:
    #     URL=URL[0:URL.find('site')]
    # except ValueError:
    #     print("substring not found")
    # https: // nioss / ncobject.jsp?id = 9153410169413622796 & site = 6032252332013374327
    # 9153410169413622796
    for n in Freq[::-1]:
        if not n  in ['0','.']:
            return ID
        else:
            if n==".":
                ID = ID[0:len(ID) - 1]
                return ID
            ID = ID[0:len(ID) - 1]
    return ID


# функция которая логинится в НИОСС
def NiossLogin(driver,login,passw):
    driver.get("https://nioss/")
    try:
        login_el=driver.find_element_by_id('user')
        login_el.send_keys(login)
        passw_el=driver.find_element_by_id('pass')
        passw_el.send_keys(passw)
        ok_button=driver.find_element_by_id('loginButton')
        ok_button.click()
    except:
        print("Логиниться не нужно")

# функция, которая преобразует номер РРН в номер площадки
def RRNtoPL(RRN):
    RRNlist=RRN.split('-')
    if len(RRNlist)>3:
        return "PL_"+RRNlist[1]+'_'+RRNlist[2]
    else:
        RRNlist = RRN.split('_')
        if len(RRNlist) > 3:
            return "PL_" + RRNlist[1] + '_' + RRNlist[2]
        else:
            return "error"

# функция, удаляет нули у айпишника
def stripIPadress(ip):
    iplist=ip.split('.')
    if len(iplist)>3:
        for item,value in enumerate(iplist):
            while iplist[item][0] == "0":
                iplist[item] = iplist[item][1:]
        return '.'.join(iplist)
    else:
        return "error"

# функция, которая открывает РРЛ по обжект айди
def NiossOpenRRL(driver,object_ID):
    time.sleep(1.5)
    driver.get("https://nioss/ncobject.jsp?id=" + object_ID+'&tab=_Parameters&mode=1')
    time.sleep(1.5)

def NiossEditRRL(driver,RRL_ID):
    try:
        NiossOpenRRL(driver,RRL_ID)
        ok_button=driver.find_element_by_id('pcEdit')
        ok_button.click()
    except Exception as e:
        print("Не удалось открыть тракт для редактирования \n убедитесь, что у вас открыт IE на нужном тракте")
        print(e)
        time.sleep(5)
        driver.quit()


#отправляем значения в браузер
@retry(WebDriverException, tries=3, delay=1.3)
def set_value_without_clear(driver,id,value):
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, id))
        )
    except:
        driver.refresh()
    if value:
        nioss_el=driver.find_element_by_id(id)
        # nioss_el.click()
        nioss_el.send_keys(value)
        time.sleep(.7)
        nioss_el.send_keys(Keys.ENTER)

#отправляем значения в браузер
@retry(WebDriverException, tries=3, delay=1.3)
def set_value(driver,id,value):
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, id))
        )
    except:
        driver.refresh()
    if value:
        nioss_el=driver.find_element_by_id(id)
        # nioss_el.click()
        nioss_el.clear()
        nioss_el.send_keys(value)
        time.sleep(.7)
        nioss_el.send_keys(Keys.ENTER)

@retry(WebDriverException, tries=3, delay=1.3)
def set_value_without_enter(driver,id,value):
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, id))
    )
    if value:
        nioss_el=driver.find_element_by_id(id)
        # nioss_el.click()
        nioss_el.clear()
        nioss_el.send_keys(value)
        time.sleep(.7)

#отправляем значения в браузер
@retry(WebDriverException, tries=3, delay=1.3)
def set_text(driver,id,name,value):
    nioss_el=driver.find_element_by_id(id).find_element_by_name(name)
    # nioss_el.click()
    nioss_el.clear()
    nioss_el.send_keys(value)
    time.sleep(.5)
    nioss_el.send_keys(Keys.ENTER)


#отправляем значения в браузер в поле селект
def sel_value(driver,id,value):
    select=Select(driver.find_element_by_id(id))
    select.select_by_visible_text(value)

@retry(WebDriverException, tries=3, delay=1.5)
def SetIPtoRRN(driver,IP,href):
    time.sleep(5)
    driver.get(href)
    RRSID = getObgectID(href,'id')
    IPlist = IP.split('.')
    driver.find_element_by_id('pcEdit').click()
    set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_0', IPlist[0])
    set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_3', IPlist[1])
    set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_6', IPlist[2])
    set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_9', IPlist[3])
    driver.find_element_by_id('theform_update').click()

def PasteParametrsEXT(driver, RRL):
    if RRL.EXT_A:
        time.sleep(5)
        driver.get(RRL.EXT_A_Href)
        RRSID = getObgectID(RRL.EXT_A_Href, 'id')
        IPlist = RRL.EXT_A_ip.split('.')
        driver.find_element_by_id('pcEdit').click()
        set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_0', IPlist[0])
        set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_3', IPlist[1])
        set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_6', IPlist[2])
        set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_9', IPlist[3])
        set_value_without_enter(driver,'id__9_9153609062913313373_'+ RRSID +'_input',RRL.EXT_A)
        driver.find_element_by_id('theform_update').click()
    if RRL.EXT_Z:
        time.sleep(5)
        driver.get(RRL.EXT_Z_Href)
        RRSID = getObgectID(RRL.EXT_Z_Href, 'id')
        IPlist = RRL.EXT_Z_ip.split('.')
        driver.find_element_by_id('pcEdit').click()
        set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_0', IPlist[0])
        set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_3', IPlist[1])
        set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_6', IPlist[2])
        set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_9', IPlist[3])
        set_value_without_enter(driver,'id__9_9153609062913313373_'+ RRSID +'_input',RRL.EXT_Z)
        driver.find_element_by_id('theform_update').click()

# функция, которая заполняет айпи адреса в РРС
def NiossPasteParametrsRRS(driver,HREF,RRLA,RRLZ):
    # открываем ближнюю сторону РРС
    driver.get(HREF[0])
    RRSID = getObgectID(driver.current_url.split('site')[0])
    IPlist=RRLA.Ip.split('.')
    driver.find_element_by_id('pcEdit').click()
    set_value(driver,'id__5_9136474451313754458_'+RRSID+'_0',IPlist[0])
    set_value(driver, 'id__5_9136474451313754458_' + RRSID + '_3', IPlist[1])
    set_value(driver, 'id__5_9136474451313754458_' + RRSID + '_6', IPlist[2])
    set_value(driver, 'id__5_9136474451313754458_' + RRSID + '_9', IPlist[3])
    set_value(driver, 'id__3_8061151301013492812_' + RRSID , RRLA.RRS.PowerRX)
    set_value(driver, 'id__3_8061151301013492811_' + RRSID, RRLA.RRS.PowerTX)
    driver.find_element_by_id('pcUpdate').click()

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9133816070113745923')))

    SetIPtoRRN(driver,RRLA.Ip,HREF[1])

    # теперь дальнюю

    driver.get(HREF[2])
    RRSID = getObgectID(driver.current_url.split('site')[0])
    IPlist = RRLA.RRS.OppositeIp.split('.')
    driver.find_element_by_id('pcEdit').click()
    set_value(driver, 'id__5_9136474451313754458_' + RRSID + '_0', IPlist[0])
    set_value(driver, 'id__5_9136474451313754458_' + RRSID + '_3', IPlist[1])
    set_value(driver, 'id__5_9136474451313754458_' + RRSID + '_6', IPlist[2])
    set_value(driver, 'id__5_9136474451313754458_' + RRSID + '_9', IPlist[3])
    set_value(driver, 'id__3_8061151301013492812_' + RRSID, RRLZ.RRS.PowerRX)
    set_value(driver, 'id__3_8061151301013492811_' + RRSID, RRLZ.RRS.PowerTX )
    driver.find_element_by_id('pcUpdate').click()

    SetIPtoRRN(driver, RRLA.RRS.OppositeIp,HREF[3])


# функция, которая открывает РРЛ по обжект айди
def NiossPasteParametrs(driver,object_ID,RRLA,RRLZ):
    dicTypeRRL={'OmniBAS-2W 16E1':'OmniBAS 2W','iPasolink EX adv':'iPasolink EX adv','UL-GX80':'ULTRALINK GX80','UL-FX80-CN':'ULTRALINK FX80 Compact Node','UL-FX80-V2':'ULTRALINK FX80','ULINK-FX80-V2':'ULTRALINK FX80','OmniBAS-8W':'OmniBAS-8W','OmniBAS-4W v2':'OmniBAS-4Wv2','OmniBAS-4Wv2':'OmniBAS-4Wv2','iPasolink 1000':'iPASOLINK 1000 IP','iPasolink 200':'iPASOLINK 200 IP','iPasolink 400':'iPASOLINK 400 IP','AMM 2p B':'Mini-Link TN','AMM 6p C':'Mini-Link TN','STRN-PTP-6250':'StreetNode V60-PTP','iPasolink EX':'iPASOLINK EX','iPasolink NEO ST':'Pasolink NEO','iPasolink NEO HP':'PASOLINK NEO HP AMR','OMNIBAS-2WCX':'OmniBas-2Wcx'}
    dicHiLow={'Upper':'Высокий','Lower':'Низкий'}
    dicReserv = {'1+0': '1+0','XPIC':'1х(1+0) XPIC','Single': '1+0'}
    dicPolar = {'1+0': 'Single','XPIC':'X-Поляризация'}

    set_value(driver,'id__9_7020954535013794924_' + object_ID +'_input',dicTypeRRL[RRLA.TypeIDU])
    time.sleep(1)
    set_value(driver,'id__9_9131240310713090205_' + object_ID + '_input', dicTypeRRL[RRLZ.TypeIDU])
    time.sleep(1)
    set_value(driver,'id__9_6061433109013797256_' + object_ID+'_input', RRLA.RRS.Band)
    set_value(driver, 'id__9_5111039577013897377_' + object_ID + '_input', RRLA.RRS.FreqTX)
    set_value(driver, 'id__9_5111039577013897378_' + object_ID + '_input', RRLA.RRS.FreqRX)
    sel_value(driver, 'id__7_9136614742213755416_' + object_ID , RRLA.RRS.Channel_Spasing)
    sel_value(driver, 'id__7_6052264868013681183_' + object_ID ,dicHiLow[RRLA.RRS.Upper])
    sel_value(driver, 'id__7_3041058345013556885_' + object_ID, dicReserv[RRLA.RRS.Reserv])
    sel_value(driver, 'id__7_9129953101913823136_' + object_ID, RRLA.RRS.AmrMax)
    # set_value(driver, 'id__3_5111039577013897372_' + object_ID , RRLA.RRS.PowerRX)
    # set_value(driver, 'id__3_5110966108013897105_' + object_ID, RRLA.RRS.PowerTX)
    # set_value(driver, 'id__3_5111039577013897373_' + object_ID, RRLZ.RRS.PowerRX)
    # set_value(driver, 'id__3_5110966108013897108_' + object_ID, RRLZ.RRS.PowerTX )
    set_value(driver, 'id__3_9130186887613885021_' + object_ID, RRLZ.RRS.Capacity_max)
    set_value(driver, 'id__0_5111039577013897389_' + object_ID, RRLZ.RRS.SubBand)
    set_text(driver, 'vv_-2' , 'common_descr',RRLA.TypeIDU + " "+RRLA.RRS.Reserv)
    # set_value(driver, 'id__9_9133804250013744569_' + object_ID+ '_input', dicPolar[RRLA.RRS.Reserv])

    ok_button = driver.find_element_by_id('pcUpdate')

    ok_button.click()

def alert_acsept(driver):
    try:
        alert = driver.switch_to_alert()
        alert.accept()
        driver.switch_to_default_content()
    except:
        print("Без уведомлений")

# функция, которая считывает параметры РРЛ из НИОСС
def NiossGetParametrs(driver,RRL_ID):
    NiossOpenRRL(driver,RRL_ID)
    RRL=NIOSS_RRL
    RRL.PL_A=driver.find_element_by_id('vv_5022250305013888246').text
    RRL.PL_B = driver.find_element_by_id('vv_5022250305013888247').text

    # edit=False
    # if not RRL.H1_A.strip(' '):
    #     print('Не заполнена высота подвеса А')
    #     RRL.H1_A=input()
    #     edit=True
    # if not RRL.H1_B.strip(' '):
    #     print('Не заполнена высота подвеса Z')
    #     RRL.H1_B=input()
    #     edit = True
    # if not RRL.D1_A.strip(' '):
    #     print('Не заполнен диаметр А')
    #     RRL.D1_A = input()
    #     edit = True
    # if not RRL.D1_B.strip(' '):
    #     print('Не заполнен диаметр Z')
    #     RRL.D1_B = input()
    #     edit = True
    # if edit:
    #     ok_button = driver.find_element_by_id('pcEdit')
    #     ok_button.click()
    #     # set_value(driver, 'id__3_9133821080013746594_' + RRL_ID,RRL.H1_A )
        # set_value(driver, 'id__3_9133821080013746595_' + RRL_ID, RRL.H1_B)
        # sel_value(driver, 'id__7_9133814060013745757_' + RRL_ID, RRL.D1_A)
        # sel_value(driver, 'id__7_9133814060013745758_' + RRL_ID, RRL.D1_B)
        # ok_button = driver.find_element_by_id('pcUpdate')
        # ok_button.click()
        # alert_acsept(driver)
    # RRL.dist = driver.find_element_by_id('vv_9133105610013559472').text
    return RRL


def NiossGetIPFromRRL(driver,HREF,HREF_RRL):
    # try:
        # открываем ближнюю сторону РРС
        driver.get(HREF[0])
        # считываем айпи адрес, если он есть
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9136474451313754458')))
        time.sleep(1.5)
        ipnear=stripIPadress(driver.find_element_by_id('vv_9136474451313754458').text)
        if ipnear=='error':
            try:
                driver.get(HREF[1])
            except:
                print("На стороне РРС А не привязан RRN. Пивяжите и ")
                input("нажмите любую клавишу для продолжения")
                NiossGetIPFromRRL(driver, HREF, HREF_RRL)
                exit()
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9136474451313754458')))
            time.sleep(1.5)
            ipnear = stripIPadress(driver.find_element_by_id('vv_9136474451313754458').text)
            if ipnear == 'error':
                driver.get(HREF_RRL)
                driver.get(HREF[2])
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9136474451313754458')))
                time.sleep(1.5)
                ipnear = stripIPadress(driver.find_element_by_id('vv_9136474451313754458').text)
                if ipnear == 'error':
                    driver.get(HREF_RRL)
                    try:
                        driver.get(HREF[3])
                    except:
                        print("На стороне РРС Z не привязан RRN. Пивяжите и ")
                        input("нажмите любую клавишу для продолжения")
                        NiossGetIPFromRRL(driver, HREF, HREF_RRL)
                        exit()
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9136474451313754458')))
                    time.sleep(1.5)
                    ipnear = stripIPadress(driver.find_element_by_id('vv_9136474451313754458').text)
                    if ipnear == 'error':
                        return input("Не нашлось ip адреса Введите вручную\n")
        driver.get(HREF_RRL)
        return ipnear
    # except Exception as e:
    #     input('Не удается определить ID RRL, введите вручную')
    #     return 'error'

def get_HREF_RRS_RRN(driver,RRL_HREF):
    try:
        NiossOpenRRL(driver, RRL_HREF)
        time.sleep(1.5)
        # открываем ближнюю сторону РРС
        RRS1_HREF=driver.find_element_by_id('vv_6071441931013972499').find_element_by_tag_name("a").get_attribute('href')
        RRS2_HREF = driver.find_element_by_id('vv_6071450625013973411').find_element_by_tag_name("a").get_attribute('href')
        driver.get(RRS1_HREF)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9133816070113745923')))
        time.sleep(2.5)
        RRN1_HREF = driver.find_element_by_id('vv_9133816070113745923').find_element_by_tag_name("a").get_attribute(
            'href')
        driver.get(RRS2_HREF)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9133816070113745923')))
        time.sleep(2.5)
        RRN2_HREF = driver.find_element_by_id('vv_9133816070113745923').find_element_by_tag_name("a").get_attribute(
            'href')
        return RRS1_HREF,RRN1_HREF,RRS2_HREF,RRN2_HREF
    except Exception as e:
        print('Проверьте привязку РРУ к РРС')
        return 'error','error','error','error'



# функция, которая определяет, нужно ле реверсировать данные при занесении
def ReversRRL(NearRRL,RRL_NIOSS):
    if NearRRL.RRS.NameAsPL == RRL_NIOSS.PL_A:
        return False
    elif NearRRL.RRS.NameAsPL == RRL_NIOSS.PL_B:
        return True
    else:
        return 'error'
def printRRL(RRL):
    print(RRL.__dict__)
    print(RRL.RRS.__dict__)
    # print(RRL.TypeIDU)
    # print(RRL.RRS.FreqTX, RRL.RRS.FreqRX, RRL.RRS.Band)
    # print(RRL.RRS.PowerTX, RRL.RRS.PowerRX)
    # print(RRL.RRS.Channel_Spasing, RRL.RRS.OppositeIp, RRL.RRS.Upper, RRL.RRS.NameAsPL)


def Add_AGOS(driver,HREF):
    # открываем ближнюю сторону РРС
    driver.get(HREF[0])

        # # считываем айпи адрес, если он есть
        # WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9136474451313754458')))
        # time.sleep(1.5)
        # ipnear=stripIPadress(driver.find_element_by_id('vv_9136474451313754458').text)
        # if ipnear=='error':
        #     driver.get(HREF[1])
        #     WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9136474451313754458')))
        #     time.sleep(1.5)
        #     ipnear = stripIPadress(driver.find_element_by_id('vv_9136474451313754458').text)
        #     if ipnear == 'error':
        #         driver.get(HREF_RRL)
        #         driver.get(HREF[2])
        #         WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9136474451313754458')))
        #         time.sleep(1.5)
        #         ipnear = stripIPadress(driver.find_element_by_id('vv_9136474451313754458').text)
        #         if ipnear == 'error':
        #             driver.get(HREF_RRL)
        #             driver.get(HREF[3])
        #             WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9136474451313754458')))
        #             time.sleep(1.5)
        #             ipnear = stripIPadress(driver.find_element_by_id('vv_9136474451313754458').text)
        #             if ipnear == 'error':
        #                 return input("Не нашлось ip адреса Введите вручную\n")
        # driver.get(HREF_RRL)
