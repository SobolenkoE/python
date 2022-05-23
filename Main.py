import RRL_Class
import SNMP
import NIOSS
import time
import Remedy_Web_1_0
import Main_Form
from tkinter import *
import AutoDO
from selenium import webdriver


def RRL_parametrs_set_to_NIOSS(driver,HREF,RRL_ID,ex):


    NearRRL = RRL_Class.RRL()
    FarRRL = RRL_Class.RRL()
    if '.' in ex.get_ip_RRL():
        NearRRL.Ip=ex.get_ip_RRL().replace("\n", "")
    else:
        NearRRL.Ip = NIOSS.NiossGetIPFromRRL(driver, HREF, 'https://nioss/ncobject.jsp?id=' + RRL_ID)

    RRL_NIOSS = NIOSS.NiossGetParametrs(driver,RRL_ID)

    NIOSS.NiossEditRRL(driver,RRL_ID)
    NearRRL.Snmp_Community = SNMP.get_community(NearRRL.Ip)
    NearRRL.getTypeIDU()
    NearRRL.getRrsProperties(RRL_NIOSS)
    FarRRL.Snmp_Community = NearRRL.Snmp_Community
    FarRRL.Ip = NearRRL.RRS.OppositeIp
    FarRRL.TypeIDU = SNMP.get_typeRRL(FarRRL.Ip,FarRRL.Snmp_Community)
    FarRRL.getRrsProperties(RRL_NIOSS)
    NIOSS.printRRL(NearRRL)
    NIOSS.printRRL(FarRRL)
    NIOSS.ReversRRL(NearRRL, RRL_NIOSS)
    if NIOSS.ReversRRL(NearRRL, RRL_NIOSS) == 'error':
        print('IP адрес не соответсвует. Или аппаратное наименовение РРС неверно.')
    elif NIOSS.ReversRRL(NearRRL, RRL_NIOSS):
        NIOSS.NiossPasteParametrs(driver, RRL_ID, FarRRL, NearRRL)
        NIOSS.NiossPasteParametrsRRS(driver, HREF, FarRRL,NearRRL)
    else:
        NIOSS.NiossPasteParametrs(driver, RRL_ID, NearRRL, FarRRL)
        NIOSS.NiossPasteParametrsRRS(driver, HREF, NearRRL,FarRRL)
    return NearRRL,FarRRL
# def RRL_

def check_d_H_RRS(driver,RRS_ID):
    NIOSS.NiossOpenRRL(driver,RRS_ID)
    RRN=driver.find_element_by_id('vv_9133816070113745923').text
    H1=driver.find_element_by_id('vv_9133816070113745934').text
    D1=driver.find_element_by_id('vv_9133814060013745752').text
    edit=False
    if not H1.strip(' '):
        print('Не заполнена высота подвеса на РРН %s' %RRN)
        H1=input()
        edit=True

    if not D1.strip(' '):
        print('Не заполнен диаметр на РРН %s' %RRN)
        D1 = input()
        edit = True
    if edit:
        ok_button = driver.find_element_by_id('pcEdit')
        ok_button.click()
        NIOSS.set_value(driver, 'id__3_9133816070113745934_' + RRS_ID,H1)
        NIOSS.sel_value(driver, 'id__7_9133814060013745752_' + RRS_ID,D1)
        # sel_value(driver, 'id__7_9133814060013745757_' + RRL_ID, RRL.D1_A)
        # sel_value(driver, 'id__7_9133814060013745758_' + RRL_ID, RRL.D1_B)
        ok_button = driver.find_element_by_id('pcUpdate')
        ok_button.click()
        # alert_acsept(driver)

    # return RRL


def RRL_NIOSS(self):

    start_time = time.time()


    options = webdriver.ChromeOptions()
    # options.add_argument('headless') #если хотим запустить chrome недивимкой
    options.add_argument("window-size=1920x1080")

    # получаем ID RRL из IE или из формы
    # try:
    RRL_ID = NIOSS.getObgectID(ex.get_ID_RRL())
    # except:
    #     RRL_ID = NIOSS.getObgectID(AutoWin.GetUrl())
    # Открываем теперь ссылку в Chrome

    driver = webdriver.Chrome(options=options)
    Credentials = Remedy_Web_1_0.load_obj('Cred')
    NIOSS.NiossLogin(driver, Credentials['login'], Credentials['pass'])

    HREF = []
    while ('' in HREF) or len(HREF)<4:
        HREF = RRS1_HREF, RRN1_HREF, RRS2_HREF, RRN2_HREF = NIOSS.get_HREF_RRS_RRN(driver,RRL_ID)
        if ('' in HREF) or len(HREF)<4:
            print("Убедитесь, что все РРН подвязаны и нажмите любую клавишу")

    check_d_H_RRS(driver,NIOSS.getObgectID(HREF[0]))
    check_d_H_RRS(driver,NIOSS.getObgectID(HREF[2]))

    # считываем все параметры с РРЛ и заносим в НИОСС.
    NearRRL,FarRRL=RRL_parametrs_set_to_NIOSS(driver,HREF,RRL_ID,ex)



    if ex.isAGOS:# Завести АГОСы
        AutoDO.RRL_add_AGOS(driver,HREF)
    if ex.isCheckList:# завести работу на проверку чек-листа
        Remedy_Web_1_0.remedyLogin(driver)
        Remedy_Web_1_0.remedy_set_param(driver, typeW='Планово-профилактическая',
                         typeS='Данные + Голос',
                         impactS='Нет влияния на сервис абонентам в зоне действия СЭ.',
                         addInfo="https://nioss/ncobject.jsp?id=" +RRL_ID,
                         supervisor='Оперативный дежурный ООКУ СРД ФО',
                         initiator='',
                         BS=NearRRL.RRS.NameAsPL[3:] + "<>" + FarRRL.RRS.NameAsPL[3:],
                         describeW='Проверка пролета по чеклисту ' + NearRRL.RRS.NameAsPL[3:] + "<>" + FarRRL.RRS.NameAsPL[3:]+'. Тип оборудования:'+ NearRRL.TypeIDU,
                         performer='', performerG='ЮГ\ЕЦУС\УИТС\MBH Юго-Запад',
                         typeNE='РРС',klassificator=['14','02','05'],incident='',
                         days=5)


    if ex.isCheckAGOS:# завести работу на проверку АГОСа
        driver.refresh()
        Remedy_Web_1_0.remedyLogin(driver)
        Remedy_Web_1_0.remedy_set_param(driver, typeW='Планово-профилактическая',
                         typeS='Данные + Голос',
                         impactS='Нет влияния на сервис абонентам в зоне действия СЭ.',
                         addInfo='',
                         supervisor='Оперативный дежурный ООКУ СРД ФО',
                         initiator='',
                         BS=NearRRL.RRS.NameAsPL[3:] + "<>" + FarRRL.RRS.NameAsPL[3:],
                         describeW='Прошу проверить АГОС по РРЛ ' + NearRRL.RRS.NameAsPL[3:] + "<>" + FarRRL.RRS.NameAsPL[3:]+'. Тип оборудования:'+ NearRRL.TypeIDU,
                         performer='', performerG='ЮГ\Краснодар\ЭТС\TN',
                         typeNE='РРС',klassificator=['15', '04', '07'], incident='',
                         days=9)

    driver.quit()
    print("Время выполнения программы: %s секунд" % (time.time() - start_time))


if __name__ == "__main__":
    root = Tk()
    ex = Main_Form.W_NIOSS_RRL(root)
    ex.go_button.bind("<Button-1>", RRL_NIOSS)

    root.mainloop()
    # RRL_NIOSS()




