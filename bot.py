
from selenium import webdriver
import NIOSS
import Remedy_Web_1_0
import AutoDO
import sys
import pickle
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


class BOT_RRL:
    def __init__(self):
        self.site_PL_A = ""
        self.site_PL_Z = ""
        self.site_A=""
        self.site_Z=""
        self.H_A=""
        self.H_Z=""
        self.D_A="0.3"
        self.D_Z="0.3"
        self.TX_freq_A=""
        self.TX_freq_Z=""
        self.TX_A=""
        self.TX_Z=""
        self.RX_A=""
        self.RX_Z=""
        self.Type_RRL_A=""
        self.Type_RRL_Z=""
        self.ip_A=""
        self.ip_Z=""
        self.Band=""
        self.Subband=""
        self.Ch_spasing=""
        self.Rezerv="1+0"
        self.Polarization="V"
        self.EXT_A=""
        self.EXT_Z=""
        self.EXT_A_ip = ""
        self.EXT_Z_ip = ""
        self.IDU_A = ""
        self.IDU_Z = ""
        self.IDU_A_exist=""
        self.IDU_Z_exist=""
        self.IDU_A_exist_name = ""
        self.IDU_Z_exist_name = ""
        self.IDU_A_exist_bool = False
        self.IDU_Z_exist_bool = False
        self.RRN_A_exist_bool = False
        self.RRN_Z_exist_bool = False
        self.EXT_A_exist_bool = False
        self.EXT_Z_exist_bool = False
        self.Upper_A = ""
        self.RRN_A_Href = ""
        self.RRN_Z_Href = ""
        self.RRS_A_Href = ""
        self.RRS_Z_Href = ""
        self.RRL_Href = ""
        self.EXT_A_Href = ""
        self.EXT_Z_Href = ""
        self.PL_A_Href = ""
        self.PL_Z_Href = ""
        self.RRN_A_name = ""
        self.RRN_Z_name = ""
        self.RRS_A_name = ""
        self.RRS_Z_name = ""
        self.RRL_name = ""
        self.EXT_A_name = ""
        self.EXT_Z_name = ""
        self.List_NE_A = []
        self.List_NE_Z = []

    def str_Z(self):
        text= 'Site_Z:=' + self.site_Z + '\n  ip:=' + self.ip_Z + '\n  Высота подвеса:=' + self.H_Z + '\n  Диаметр:=' + self.D_Z
        text += '\n  Частота:=' + self.TX_freq_Z + '\n  Мощность передатчика:=' + self.TX_Z + '\n  Уровень на приёме:=' + self.RX_Z + '\n  Тип внутреннего блока:=' + self.IDU_Z + '\n'
        return text

    def str_A(self):
        text = '\nSite_A:=' + self.site_A
        text += '\n  ip:=' + self.ip_A + '\n  Тип РРЛ:=' + self.Type_RRL_A + '\n  Высота подвеса:=' + self.H_A + '\n  Диаметр:=' + self.D_A
        text += '\n  Частота:=' + self.TX_freq_A + '\n  Мощность передатчика:=' + self.TX_A + '\n  Уровень на приёме:=' + self.RX_A + '\n  Тип внутреннего блока:=' + self.IDU_A + '\n'
        return text

    def str_ext(self):
        text = '\n  Поляризация:=' + self.Polarization + '\n  Резервирование:=' + self.Rezerv + '\n  Ширина спектра:=' + self.Ch_spasing + '\n_____'
        return text

    def __str__(self):
        text=self.str_A() + self.str_Z()+ self.str_ext()
        return text



def get_driver():
    options=webdriver.ChromeOptions()
    options.add_argument('--allow-profiles-outside-user-dir')
    options.add_argument('--enable-profile-shortcut-manager')
    # options.add_argument('user-data-dir=C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user')
    # options.add_argument('--profile-directory=C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user\\Profile 1')
    # options.add_argument('--profile-directory=Default')
    options.add_argument('--disable-gpu')

    # options.add_argument("--disable-notifications")
    # if '!' not in sys.argv[3].lower():
    # options.add_argument('headless') #если хотим запустить chrome недивимкой
    options.add_argument("window-size=1220x1580")
    options.add_experimental_option('useAutomationExtension', False)

    try:

        driver = webdriver.Chrome(options=options)

        return driver


    except:
        time.sleep(30)



def getDigitsInt(text):
    # https://nioss/ncobject.jsp?id=9153410135213621294&site=6031643620013350054
    # https:/nioss/ncobject.jsp?id=9153410135213621294
    ID=''
    for n in text[::-1]:
        if n.isdigit():
            ID=n+ID
    if not ID:
        ID="0"
    return ID

def fill_form():
    text = 'Site_A:=23\n  ip:=\n  Тип РРЛ:=\n  Высота подвеса:=\n  Диаметр:=0.3\n  Частота:=\n  Мощность передатчика:=\n  Уровень на приёме:=\n  Тип внутреннего блока:=\n'
    text += 'Site_Z:=23\n  ip:=\n  Высота подвеса:=\n  Диаметр:=0.3\n  Частота:=\n  Мощность передатчика:=\n  Уровень на приёме:=\n  Тип внутреннего блока:=\n'
    text += '\n  Поляризация:=V\n  Резервирование:=1+0\n  Ширина спектра:=\n_____'
    return text

def reply_send(to_address,subject,text,att_path=''):
    # инициализируем объект outlook
    olk = win32com.client.Dispatch("Outlook.Application")
    Msg = olk.CreateItem(0)
    try:
        performer = Remedy_Web_1_0.get_performer_by_mail(to_address,'TN')[0]
    except:
        print('Ошибка распознавания в письме')
        Remedy_Web_1_0.reply_send(to_address, 'Пользователь', 'Автозаведение работ', 'Ошибка распознавания текста в письме!')
        sys.exit()
    # формируем письма, выставляя адресата, тему и текст
    Msg.To = to_address

    Msg.Subject =  subject  # добавляем RE в тему
    if att_path:
        Msg.Attachments.Add(att_path)
    Msg.Body = "Здравствуйте, " + str(performer) + "! \n" \
                                                 +text

        # и отправляем
    Msg.Send()

def NiossPasteParametrs(driver,RRL):
    # dicTypeRRL={'RTN 380AX':'RTN 380AX','OmniBAS 2W':'OmniBAS 2W','iPasolink EX adv':'iPasolink EX adv','UL-GX80':'ULTRALINK GX80','UL-FX80-CN':'ULTRALINK FX80 Compact Node','UL-FX80-V2':'ULTRALINK FX80','ULINK-FX80-V2':'ULTRALINK FX80','OmniBAS-8W':'OmniBAS-8W','OmniBAS-4W v2':'OmniBAS-4Wv2','OmniBAS-4Wv2':'OmniBAS-4Wv2','iPasolink 1000':'iPASOLINK 1000 IP','iPasolink 200':'iPASOLINK 200 IP','iPasolink 400':'iPASOLINK 400 IP','AMM 2p B':'Mini-Link TN','AMM 6p C':'Mini-Link TN','STRN-PTP-6250':'StreetNode V60-PTP','iPasolink EX':'iPASOLINK EX','iPasolink NEO ST':'Pasolink NEO','iPasolink NEO HP':'PASOLINK NEO HP AMR','OMNIBAS-2WCX':'OmniBas-2Wcx'}
    dicHiLow={'Upper':'Высокий','Lower':'Низкий'}
    # dicReserv = {'1+0': '1+0','XPIC':'1х(1+0) XPIC','Single': '1+0'}
    # dicPolar = {'1+0': 'Single','XPIC':'X-Поляризация'}
    time.sleep(5)
    driver.get(RRL.RRL_Href)
    object_ID = NIOSS.getObgectID(RRL.RRL_Href,'id')
    try:
        driver.find_element_by_id('pcEdit').click()
    except:
        pass
    # id__9_7020954535013794924_9162654882313532531_input
    NIOSS.set_value(driver,'id__9_7020954535013794924_' + object_ID +'_input',RRL.Type_RRL_A)
    time.sleep(2)
    NIOSS.set_value(driver,'id__9_9131240310713090205_' + object_ID + '_input', RRL.Type_RRL_Z)
    time.sleep(2)
    NIOSS.set_value(driver,'id__9_6061433109013797256_' + object_ID+'_input', RRL.Band)
    NIOSS.set_value(driver, 'id__9_5111039577013897377_' + object_ID + '_input', RRL.TX_freq_A)
    time.sleep(10)
    NIOSS.set_value(driver, 'id__9_5111039577013897378_' + object_ID + '_input', RRL.TX_freq_Z)
    time.sleep(10)
    NIOSS.sel_value(driver, 'id__7_9136614742213755416_' + object_ID , RRL.Ch_spasing)
    NIOSS.sel_value(driver, 'id__7_6052264868013681183_' + object_ID ,dicHiLow[RRL.Upper_A])
    NIOSS.sel_value(driver, 'id__7_3041058345013556885_' + object_ID, RRL.Rezerv)
    driver.find_element_by_id('id__0_6050356291013956598_' + object_ID ).send_keys(Keys.ENTER)

    # NIOSS.sel_value(driver, 'id__7_9129953101913823136_' + object_ID, RRLA.RRS.AmrMax)
    # set_value(driver, 'id__3_5111039577013897372_' + object_ID , RRLA.RRS.PowerRX)
    # set_value(driver, 'id__3_5110966108013897105_' + object_ID, RRLA.RRS.PowerTX)
    # set_value(driver, 'id__3_5111039577013897373_' + object_ID, RRLZ.RRS.PowerRX)
    # set_value(driver, 'id__3_5110966108013897108_' + object_ID, RRLZ.RRS.PowerTX )
    # set_value(driver, 'id__3_9130186887613885021_' + object_ID, RRLZ.RRS.Capacity_max)
    # set_value(driver, 'id__0_5111039577013897389_' + object_ID, RRLZ.RRS.SubBand)
    # NIOSS.set_text(driver, 'vv_-2' , 'common_descr',RRLA.TypeIDU + " "+RRLA.RRS.Reserv)
    # set_value(driver, 'id__9_9133804250013744569_' + object_ID+ '_input', dicPolar[RRLA.RRS.Reserv])
    # time.sleep(5)
    # driver.find_element_by_id('theform_update').click()




def NiossPasteParametrsRRS(driver,RRL):
    # открываем ближнюю сторону РРС
    time.sleep(5)
    driver.get(RRL.RRS_A_Href)
    RRSID = NIOSS.getObgectID(RRL.RRS_A_Href,'id')
    IPlist=RRL.ip_A.split('.')
    driver.find_element_by_id('pcEdit').click()
    NIOSS.set_value_without_enter(driver,'id__5_9136474451313754458_'+RRSID+'_0',IPlist[0])
    NIOSS.set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_3', IPlist[1])
    NIOSS.set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_6', IPlist[2])
    NIOSS.set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_9', IPlist[3])
    NIOSS.set_value_without_enter(driver, 'id__3_8061151301013492812_' + RRSID , RRL.RX_A)
    NIOSS.set_value_without_enter(driver, 'id__3_8061151301013492811_' + RRSID, RRL.TX_A)
    NIOSS.sel_value(driver, 'id__7_9133814060013745752_' + RRSID, RRL.D_A)
    NIOSS.sel_value(driver, 'id__7_9157348607313279670_' + RRSID, RRL.Polarization)
    NIOSS.set_value(driver, 'id__3_9133816070113745934_' + RRSID, RRL.H_A)

    try:
        driver.find_element_by_id('theform_update').click()
    except:
        pass

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'vv_9133816070113745923')))


    # теперь дальнюю
    time.sleep(5)
    driver.get(RRL.RRS_Z_Href)
    RRSID = NIOSS.getObgectID(RRL.RRS_Z_Href,'id')
    IPlist = RRL.ip_Z.split('.')
    driver.find_element_by_id('pcEdit').click()
    NIOSS.set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_0', IPlist[0])
    NIOSS.set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_3', IPlist[1])
    NIOSS.set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_6', IPlist[2])
    NIOSS.set_value_without_enter(driver, 'id__5_9136474451313754458_' + RRSID + '_9', IPlist[3])
    NIOSS.set_value_without_enter(driver, 'id__3_8061151301013492812_' + RRSID, RRL.RX_Z)
    NIOSS.set_value_without_enter(driver, 'id__3_8061151301013492811_' + RRSID, RRL.TX_Z )
    NIOSS.sel_value(driver, 'id__7_9133814060013745752_' + RRSID, RRL.D_Z)
    NIOSS.sel_value(driver, 'id__7_9157348607313279670_' + RRSID, RRL.Polarization)
    NIOSS.set_value(driver, 'id__3_9133816070113745934_' + RRSID, RRL.H_Z)
    try:
        driver.find_element_by_id('theform_update').click()
    except:
        pass

def clear_str(string):
    if string:
        while string and (not ((string[0].isalpha() or string[0].isdigit()) and (string[-1].isalpha() or string[-1].isdigit()))):
            if not (string[0].isalpha() or string[0].isdigit()):
                string=string[1:]
            if string and (not (string[-1].isalpha() or string[-1].isdigit())):
                string=string[:-1]

    return string

def validate_RRL(RRL):
    # Площадка А и Z
    error=False
    error_text=''
    RRL.site_A=getDigitsInt(RRL.site_A)
    RRL.site_Z = getDigitsInt(RRL.site_Z)
    if len(RRL.site_A)<5:
        error=True
        error_text += 'Не правильно введен номер площадки А!\n'
    if len(RRL.site_Z)<5:
        error=True
        error_text += 'Не правильно введен номер площадки Z!\n'
    RRL.site_PL_A = "PL_" + RRL.site_A[0:2] + "_" + RRL.site_A[2:]
    RRL.site_PL_Z = "PL_" + RRL.site_Z[0:2] + "_" + RRL.site_Z[2:]

    RRL.RX_A ='-'+RRL.RX_A
    RRL.RX_Z = '-' + RRL.RX_Z



    if '1+0'in RRL.Rezerv:
        if 'xpic' in RRL.Rezerv.lower():
            RRL.Rezerv='1х(1+0) XPIC'
        else:
            RRL.Rezerv = '1+0'
    elif '2' in RRL.Rezerv:
        if 'xpic' in RRL.Rezerv.lower():
            RRL.Rezerv='2х(1+0) XPIC'
        else:
            RRL.Rezerv = '2+0'
    else:
        error = True
        error_text += 'Укажите резерирование в формате 1+0, 1+0XPIC, 2xXPIC!\n'

    if any(('hv'in RRL.Polarization.lower(),'vh'in RRL.Polarization.lower(),'x'in RRL.Polarization.lower(),'х'in RRL.Polarization.lower(),('v'in RRL.Polarization.lower()) and ('h' in RRL.Polarization.lower()))):
        RRL.Polarization='V/H'
    elif 'h'in RRL.Polarization.lower():
        RRL.Polarization = 'H'
    elif 'v' in RRL.Polarization.lower():
        RRL.Polarization = 'V'
    else:
        error = True
        error_text += 'Укажите поляризацию в формате X,V,H!\n'

    if 'нет' in RRL.IDU_A_exist.lower():
        RRL.IDU_A_exist_bool=True
    if 'нет' in RRL.IDU_Z_exist.lower():
        RRL.IDU_Z_exist_bool = True
    # диапазон
    if 70000<int(RRL.TX_freq_A) <90000:RRL.Band='70/80'
    elif 36000<int(RRL.TX_freq_A) <40000:RRL.Band='38'
    elif 17000<int(RRL.TX_freq_A) <20000:RRL.Band='18'
    elif 14000<int(RRL.TX_freq_A) <17000:RRL.Band='15'
    elif 12000<int(RRL.TX_freq_A) <14000:RRL.Band='13'
    elif 10000<int(RRL.TX_freq_A) <12000:RRL.Band='11'
    elif 7800<int(RRL.TX_freq_A) <9000:RRL.Band='8'
    elif 6800<int(RRL.TX_freq_A) <7800:RRL.Band='7'
    elif 5800<int(RRL.TX_freq_A) <6800:RRL.Band='6'
    else:
        error = True
        error_text += 'Не правильно указана частота!\n'
    # верх/низ
    if int(RRL.TX_freq_Z)<int(RRL.TX_freq_A):
        RRL.Upper_A='Upper'
    else:
        RRL.Upper_A = 'Lower'
    # Тип оборудования
    # сначала определяем функции для определения типа внутреннего блока
    def selectOmni(type):
        try:
            return {'8w' in type: 'OmniBAS-8W',
                    '4w' in type: 'OmniBAS-4W',
                    '4wcx' in type: 'OmniBAS-4Wcx',
                    '2w' in type: 'OmniBAS-2W',
                    '2wcx' in type: "OmniBAS-2Wcx"}[True]  # вычисления в словаре начинаются с конца
        except:
            pass

    def selectML66XX(type):
        try:
            return {'6654' in type: 'MINI-LINK 6654',
                    '6651' in type: 'MINI-LINK 6651',
                    '6692' in type: 'MINI-LINK 6692'}[True]  # вычисления в словаре начинаются с конца
        except:
            pass
    # словарь для функций
    dict_func = {'omni': selectOmni,
                 '66': selectML66XX}

    if '380' in RRL.Type_RRL_A: RRL.Type_RRL_Z= RRL.Type_RRL_A="RTN 380AX"
    elif 'FX' in RRL.Type_RRL_A: RRL.Type_RRL_A= RRL.Type_RRL_Z="ULTRALINK FX80"
    elif 'GX' in RRL.Type_RRL_A: RRL.Type_RRL_A=RRL.Type_RRL_Z = "UlLTRALINK GX80"
    elif 'EX' in RRL.Type_RRL_A: RRL.Type_RRL_A=RRL.Type_RRL_Z = "iPASOLINK EX"
    elif ('omni' in RRL.Type_RRL_A.lower()): RRL.Type_RRL_A, RRL.Type_RRL_Z = map(dict_func['omni'],(RRL.IDU_A.lower(),RRL.IDU_Z.lower()))
    elif '66' in RRL.Type_RRL_A: RRL.Type_RRL_A, RRL.Type_RRL_Z = map(dict_func['66'],(RRL.IDU_A.lower(),RRL.IDU_Z.lower()))
    elif any(('2p' in RRL.IDU_A.lower(),'6p' in RRL.IDU_A.lower(),'20p' in RRL.IDU_A.lower())): RRL.Type_RRL_A =RRL.Type_RRL_Z = "Mini-Link TN"
    else:
        error = True
        error_text += 'Я пока не знаю такого типа оборудования!\n'


    if RRL.Band =='70/80':
        if RRL.IDU_A:
            if '905' in RRL.IDU_A: RRL.EXT_A='RTN 905 2F'
            else:
                error = True
                error_text += 'Не могу определить тип коммутатора на площадке А!\n'
            RRL.IDU_A=RRL.EXT_A
        if RRL.IDU_Z:
            if '905' in RRL.IDU_Z: RRL.EXT_Z = 'RTN 905 2F'
            else:
                error = True
                error_text += 'Не могу определить тип коммутатора на площадке Z!\n'
            RRL.IDU_Z = RRL.EXT_Z


        if (RRL.IDU_A and not RRL.EXT_A_ip) or (RRL.IDU_Z and not RRL.EXT_Z_ip):
            error = True
            error_text += 'Заполните ip коммутатора!\n'
            error_text += RRL.str_A()
            if RRL.IDU_A and not RRL.EXT_A_ip:
                error_text +='  ip EXT:=\n\n'
            error_text += RRL.str_Z()
            if RRL.IDU_Z and not RRL.EXT_Z_ip:
                error_text +='  ip EXT:=\n\n'
            error_text += RRL.str_ext()


    if error:
        return error, error_text
    else:
        return error,RRL
    # смотрим есть ли блок EXT


def extract_rrl(RRL,data_a_list,data_z_list):
    error_text=""
    error=False

    fields_A={'Site_A':'site_A',
                'ip':'ip_A',
                'Тип РРЛ':'Type_RRL_A',
                'Высота подвеса':'H_A',
                'Диаметр':'D_A',
                'Частота':'TX_freq_A',
                'Мощность передатчика':'TX_A',
                'Уровень на приёме':'RX_A',
                'Тип внутреннего блока':'IDU_A'}

    for field in fields_A:
        try:
            setattr(RRL,fields_A[field],data_a_list[field])
        except:
            error = True
            error_text += 'Не удалось считать параметр: %s со стороны А\n' %field

    fields_Z = {'Site_Z': 'site_Z',
                'ip': 'ip_Z',
                'Высота подвеса': 'H_Z',
                'Диаметр': 'D_Z',
                'Частота': 'TX_freq_Z',
                'Мощность передатчика': 'TX_Z',
                'Уровень на приёме': 'RX_Z',
                'Тип внутреннего блока': 'IDU_Z',
                'Поляризация':'Polarization',
                'Ширина спектра':'Ch_spasing',
                'Резервирование':'Rezerv'
                }

    for field in fields_Z:
        try:
            setattr(RRL, fields_Z[field], data_z_list[field])
        except:
            error = True
            error_text += 'Не удалось считать параметр: %s со стороны Z\n' % field

    fields_optional = {'Внутренний блок новый(да/нет': ('IDU_A_exist','IDU_Z_exist'),
                       'ip EXT':('EXT_A_ip','EXT_Z_ip')}

    for field in fields_optional:
        try:
            setattr(RRL, fields_optional[field][0], data_a_list[field])
        except:
            try:
                setattr(RRL, fields_optional[field][1], data_z_list[field])
            except:
                pass

    if error:
        return error, error_text
    else:
        return error,RRL


def get_data_rrl_from_mail(text,RRL):

    data_a_list = {}
    data_z_list = {}
    Data_A,Data_Z=text.split('Site')[1:3]
    Data_A='Site'+Data_A
    Data_Z = 'Site' + Data_Z
    data_split=Data_A.split('\n')


    for string in data_split:
        if ':=' in string:
            data_a_list[clear_str(string.split(':=')[0])]=clear_str(string.split(':=')[1])
    data_split = Data_Z.split('\n')

    for string in data_split:
        if ':=' in string:
            data_z_list[clear_str(string.split(':=')[0])]=clear_str(string.split(':=')[1])
    return extract_rrl(RRL,data_a_list,data_z_list)

# функция проверяет есть ли в НИОСС номера БС
# и какие сетевые элементы привязаны к каждой из БС
def get_existing_NE(driver,RRL):
    error = False
    error_text = ''
    RRL.PL_A_Href = AutoDO.check_BS_NIOSS(driver, RRL.site_PL_A)
    if not RRL.PL_A_Href:
        error = True
        error_text += 'Не могу найти ' + RRL.site_PL_A + '!\n'
    RRL.PL_Z_Href = AutoDO.check_BS_NIOSS(driver, RRL.site_PL_Z)
    if not RRL.PL_Z_Href:
        error = True
        error_text += 'Не могу найти ' + RRL.site_PL_Z + '!\n'
    if not error:
        RRL.List_NE_A = AutoDO.get_list_NE_by_bshref(driver, RRL.PL_A_Href)
        RRL.List_NE_Z = AutoDO.get_list_NE_by_bshref(driver, RRL.PL_Z_Href)
    else:
        return error,error_text

    return error, RRL


def get_existing_RRN_RRS_RRL(driver,RRL):
    RRN_exist = [x for x in RRL.List_NE_A if
                 ('RRS' in x[0] and x[4] == 'Планируется' and RRL.site_PL_Z[3:] in x[0] )]
    if RRN_exist:
        RRL.RRS_A_Href=RRN_exist[0][-1]
        RRL.RRS_A_name = RRN_exist[0][0]
        # пробуем найти РРЛ к РРС
        try:
            time.sleep(2)
            driver.get(RRL.RRS_A_Href)
            time.sleep(2)
            RRL.RRL_Href = driver.find_element_by_xpath('//*[@id="vv_7022045247013797223"]/span/span/a').get_attribute('href')
        except:
            pass
        # пробуем найти RRN и EXT к РРС
        try:
            RRNs = driver.find_elements_by_xpath('//*[@id="vv_9133816070113745923"]/span/span/a')
            if RRNs:
                for RRN in RRNs:
                    if 'EXT' in RRN.get_attribute('innerText'):
                        RRL.EXT_A_Href=RRN.get_attribute('href')
                        RRL.EXT_A_name = RRN.get_attribute('innerText')
                        RRL.EXT_A_exist_bool = True
                    else:
                        RRL.RRN_A_Href = RRN.get_attribute('href')
                        RRL.RRN_A_name = RRN.get_attribute('innerText')
                        RRL.RRN_A_exist_bool = True
        except:
            pass

    RRN_exist = [x for x in RRL.List_NE_Z if
                 ('RRS' in x[0] and x[4] == 'Планируется' and RRL.site_PL_A[3:] in x[0])]
    if RRN_exist:
        RRL.RRS_Z_Href = RRN_exist[0][-1]
        RRL.RRS_Z_name = RRN_exist[0][0]
        # пробуем найти RRN и EXT к РРС
        try:
            time.sleep(2)
            driver.get(RRL.RRS_Z_Href)
            time.sleep(2)
            RRNs = driver.find_elements_by_xpath('//*[@id="vv_9133816070113745923"]/span/span/a')
            if RRNs:
                for RRN in RRNs:
                    if 'EXT' in RRN.get_attribute('innerText'):
                        RRL.EXT_Z_Href = RRN.get_attribute('href')
                        RRL.EXT_Z_name = RRN.get_attribute('innerText')
                        RRL.EXT_Z_exist_bool = True
                    else:
                        RRL.RRN_Z_Href = RRN.get_attribute('href')
                        RRL.RRN_Z_name = RRN.get_attribute('innerText')
                        RRL.RRN_Z_exist_bool = True
        except:
            pass

    # пробуем найти пустые RRN которые можно подтянуть, если РРС еще не подтянутся к РРУ

    if not RRL.RRN_A_exist_bool:
        RRN_exist =[x for x in RRL.List_NE_A if
                    (('RRN' in x[0]) and ('EXT' not in x[0])) and ((x[3] == RRL.ip_A) or ((x[1] == '') and (x[2]=='') and (x[3]=='')))]
        if RRN_exist:
            RRL.RRN_A_Href=RRN_exist[0][-1]
            RRL.RRN_A_name = RRN_exist[0][0]
    RRN_exist = [x for x in RRL.List_NE_A if
                 ('EXT' in x[0]) and (
                         (x[3] == RRL.ip_A) or ((x[1] == '') and (x[2] == '') and (x[3] == '')))]
    if RRN_exist:
        RRL.EXT_A_Href = RRN_exist[0][-1]
        RRL.EXT_A_name = RRN_exist[0][0]

    if not RRL.RRN_Z_exist_bool:
        RRN_exist = [x for x in RRL.List_NE_Z if
                     (('RRN' in x[0]) and ('EXT' not in x[0])) and ((x[3] == RRL.ip_Z) or (x[1] == '' and x[2]=='' and x[3]==''))]
        if RRN_exist:
            RRL.RRN_Z_Href = RRN_exist[0][-1]
            RRL.RRN_Z_name = RRN_exist[0][0]
    RRN_exist = [x for x in RRL.List_NE_Z if
                 ('EXT' in x[0]) and (
                         (x[3] == RRL.ip_Z) or ((x[1] == '') and (x[2] == '') and (x[3] == '')))]

    if RRN_exist:
        RRL.EXT_Z_Href = RRN_exist[0][-1]
        RRL.EXT_Z_name = RRN_exist[0][0]

    return RRL







if __name__ == "__main__":
    pathname = 'C:\\Users\\yvsobol3\\PycharmProjects\\Test'

    # sys.argv = ['C:\\Users\\yvsobol3\\PycharmProjects\\Test','\n Site_A:=231102\n ip:=11.76.9.217\n Тип РРЛ:=OmniBAS\n Высота подвеса:=60\n Диаметр:=0.6\nВнутренний блок новый (да/нет):=нет\n Частота:=19480\n Мощность передатчика:=24\n Уровень на приёме:=-36\n Тип внутреннего блока:=OmniBAS-8W\nSite_Z:=232521\n ip:=11.76.9.220\n Высота подвеса:=30\n Диаметр:=0.6\n Частота:=18470\n Мощность передатчика:=24\nВнутренний блок новый (да/нет):=нет\n Уровень на приёме:=-36\n Тип внутреннего блока:=OmniBAS-2Wcx\n\n Поляризация:=VH\n Резервирование:=1+0 XPIC\n Ширина спектра:=56\n\n\nОтправлено из мобильного приложения Яндекс.Почты\n', 'drobotenko_s.a@mail.ru', 'Ниосс']
    # RRS_A_HREF='https://nioss/ncobject.jsp?id=9162706357013121174'
    # RRS_Z_HREF= 'https://nioss/ncobject.jsp?id=9162706365013122148'
    # RRL_HREF='https://nioss/ncobject.jsp?id=9162706378113123520'
    # RRN_A_HREF= 'https://nioss/ncobject.jsp?id=9162706370513124039'
    # RRN_Z_HREF = 'https://nioss/ncobject.jsp?id=9162706375713124969'
    # options = webdriver.ChromeOptions()
    # options.add_argument('headless') #если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1220x2080")
    # driver = webdriver.Chrome(options=options)
    # driver.get("https://nioss/")
    # sys.argv=load_obj('argv')
    save_obj(sys.argv, 'argv')
    if (('ниос' in sys.argv[3].lower()) or ('nios' in sys.argv[3].lower())) and ('Новый пролет в NIOSS bot chat' not in sys.argv[3]):
        if ('занести' in sys.argv[1].lower()) and ('данные' in sys.argv[1].lower()):
            text = 'Похоже, вы хотите занести данные в НИОСС!\nЕсли это так, то заполните данные по новой РРЛ ниже в ответном письме:\n'
            text+=fill_form()
            reply_send(sys.argv[2], 'Новый пролет в NIOSS bot chat', text)
    if (('ниос' in sys.argv[3].lower()) or ('nios' in sys.argv[3].lower())) and (('site_' in sys.argv[1].lower()) and ('высота подвеса:=' in sys.argv[1].lower())):
        RRL = BOT_RRL()
        error,RRL=get_data_rrl_from_mail(sys.argv[1],RRL)
        if error:
            text = 'Похоже, что при заведении данных возникла ошибка!\n' +RRL+ 'Заполните данные по новой РРЛ, как указано ниже, в ответном письме:\n'
            text += fill_form()
            reply_send(sys.argv[2], 'Новый пролет в NIOSS bot', text)
            sys.exit()
        error,RRL=validate_RRL(RRL)
        if error:
            text = 'Мне не все понятно!\n' + RRL + 'Исправьте или дополните данные и попробуйте заново!\n'
            reply_send(sys.argv[2], 'Новый пролет в NIOSS bot', text)
            sys.exit()
        driver = get_driver()
        # проверяем наличие площадки и собираем списки сетевых элементов на каждую сторону
        error, RRL = get_existing_NE(driver, RRL)
        if error:
            text = 'Мне не все понятно!\n' + RRL + 'Исправьте или дополните данные и попробуйте заново!\n'
            reply_send(sys.argv[2], 'Новый пролет в NIOSS bot', text)
            sys.exit()

        RRL = get_existing_RRN_RRS_RRL(driver,RRL)
        error, RRL = AutoDO.RRL_birth(driver,
                                      RRL)
        if error == 'error':
            text = 'Мне не все понятно!\n' + RRL + 'Исправьте или дополните данные и попробуйте заново!\n'
            reply_send(sys.argv[2], 'Новый пролет в NIOSS bot', text)
            sys.exit()
        if RRL.Band=='70/80':
            NIOSS.PasteParametrsEXT(driver, RRL)
        else:# здесь ситуация с классическими релейками, разбор сущ. корзина или новая
            pass



        NIOSS.SetIPtoRRN(driver, RRL.ip_A, RRL.RRN_A_Href)
        NIOSS.SetIPtoRRN(driver, RRL.ip_Z, RRL.RRN_Z_Href)
        NiossPasteParametrsRRS(driver, RRL)
        NiossPasteParametrs(driver, RRL)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        # driver.find_element_by_tag_name('body').screenshot(pathname + '\\ScreenShot ' + RRL.site_PL_A + ' - ' + RRL.site_PL_Z + '.png')
        driver.save_screenshot(pathname + '\\Scr\\ScreenShot ' + RRL.site_PL_A + ' - ' + RRL.site_PL_Z + '.png')
        text = 'Информация о пролете ' + RRL.site_PL_A + '<>' + RRL.site_PL_Z + ' занесена.'
        reply_send(sys.argv[2], 'Новый пролет в NIOSS bot chat', text,pathname + '\\Scr\\ScreenShot ' + RRL.site_PL_A + ' - ' + RRL.site_PL_Z + '.png' )
        driver.quit()
        Remedy_Web_1_0.main(['',RRL.site_A+' прошу проверить пролет ' +RRL.site_A+'<>'+RRL.site_Z+ ' ' +RRL.Type_RRL_A+' по чек листу.',sys.argv[2],'newwork уитс'])
        if RRL.EXT_A_Href:
            Remedy_Web_1_0.main(['',
                                 RRL.Site_A + ' прошу проверить EXT ' + RRL.EXT_A_name +' по чек листу.',
                                 sys.argv[2], 'newwork уитс'])
        if RRL.EXT_Z_Href:
            Remedy_Web_1_0.main(['',
                                 RRL.Site_Z + ' прошу проверить EXT ' + RRL.EXT_Z_name + ' по чек листу.',
                                 sys.argv[2], 'newwork уитс'])


    # sys.argv[1] = тело письма
    # sys.argv[3] = тема письма
