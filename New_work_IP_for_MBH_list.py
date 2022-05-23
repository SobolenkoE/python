
import time
import Remedy_Web
import Remedy_Web_1_0
import Main
import sys
from tkinter import *
# -:port=6/1/1-:PL=PL_23_2479-:index=1-:type=8603-:VLAN=500-:IP=10.220.15.153-:lo0=10.62.6.56
# -:port=6/1/2-:PL=PL_23_1290-:index=1-:type=8615-:VLAN=502-:IP=10.220.15.217-:lo0=10.62.6.57
def get_param_transport(value):
    dicVal={}
    list_param=value.split("-:")
    list_param.remove('')
    for n,v in enumerate(list_param):
        list_param[n]=v.split("=")

    dicVal=dict(list_param)
    # for v in list_param:
    #     dicVal[v.split('=')[0]] = {v.split('=')[1]}
    return "VLAN=" + dicVal['VLAN']+" IP=" + dicVal['IP']+"\r lo0=" + dicVal['lo0']+"\r port test 8630=" + dicVal['port'], "Прошу организовать транспорт к новому узлу " + dicVal['type'] + " MBH"+dicVal['PL'][2:] +"_"+dicVal['index'],"MBH"+dicVal['PL'][2:] +"_"+dicVal['index']

def get_param_pred(value):
    dicVal={}
    list_param=value.split("-:")
    list_param.remove('')
    for n,v in enumerate(list_param):
        list_param[n]=v.split("=")

    dicVal=dict(list_param)
    # for v in list_param:
    #     dicVal[v.split('=')[0]] = {v.split('=')[1]}
    return "VLAN=" + dicVal['VLAN']+" IP=" + dicVal['IP']+"\r lo0=" + dicVal['lo0']+"\r port test 8630=" + dicVal['port'], "Прошу выполнить преднастройку узла " + dicVal['type'] + " MBH"+dicVal['PL'][2:] +"_"+dicVal['index'],"MBH"+dicVal['PL'][2:] +"_"+dicVal['index']

def get_param_integ(value):
    dicVal={}
    list_param=value.split("-:")
    list_param.remove('')
    for n,v in enumerate(list_param):
        list_param[n]=v.split("=")

    dicVal=dict(list_param)
    # for v in list_param:
    #     dicVal[v.split('=')[0]] = {v.split('=')[1]}
    return "VLAN=" + dicVal['VLAN']+" IP=" + dicVal['IP']+"\r lo0=" + dicVal['lo0']+"\r port test 8630=" + dicVal['port'], "Прошу интегрировать узел в СУ " + dicVal['type'] + " MBH"+dicVal['PL'][2:] +"_"+dicVal['index'],"MBH"+dicVal['PL'][2:] +"_"+dicVal['index']


if __name__ == "__main__":
    # sys.argv=['',"",""]
    # sys.argv[0] = 'C:\\Users\\yvsobol3\\PycharmProjects\\Test\\Remedy_web.py'
    # sys.argv[1] = '-:PL=PL_23_0162-:Type=8615-:MBH=MBH_23_0162_RR2\r'
    # # sys.argv[1] = '-:port=6/1/6-:PL=PL_23_0716-:index=1-:type=8615-:VLAN=553-:IP=10.220.7.121-:lo0=10.62.6.74\r'
    #
    # sys.argv[2] ="Чеклист"

    start_time = time.time()


    incident=''
    opercontext=''


    for value in sys.argv[1].split("\r")[:-1]:
        options = Remedy_Web_1_0.webdriver.ChromeOptions()
        driver = Remedy_Web_1_0.get_driver(options)
        # Получаем словарь из Excel
        list_param = value.split("-:")
        list_param.remove('')
        for n, v in enumerate(list_param):
            list_param[n] = v.split("=")
        dicVal = dict(list_param)

        if sys.argv[2]=="Транспорт":
            Remedy_Web_1_0.remedyLogin(driver)
        # for value in str1.split("\n"):
            addInfo,describeW,BS=get_param_transport(value)
            performer=""
            performerG='ЮГ\ЕЦУС\УИТС\MBH Юго-Запад'
            BS = BS[4:-2]
            Remedy_Web_1_0.remedy_set_param(driver,typeW='Планово-профилактическая',
            typeS='Данные + Голос',
            impactS='Нет влияния на сервис абонентам в зоне действия СЭ.',
            addInfo=addInfo,
            supervisor='Оперативный дежурный ООКУ СРД ФО',
            initiator='',
            BS=BS, describeW=describeW,performer=performer, performerG=performerG,
            typeNE='MBH',klassificator=['14','04','01'],incident=incident,opercontext=opercontext,
                             days=15)

        elif sys.argv[2]=="Преднастройка":
            Remedy_Web_1_0.remedyLogin(driver)
            # for value in str1.split("\n"):
            addInfo, describeW, BS = get_param_pred(value)
            performer = ""
            performerG = 'ЮГ\ЕЦУС\УИТС\MBH Юго-Запад'
            BS = BS[4:-2]
            Remedy_Web_1_0.remedy_set_param(driver, typeW='Планово-профилактическая',
                                        typeS='Данные + Голос',
                                        impactS='Нет влияния на сервис абонентам в зоне действия СЭ.',
                                        addInfo=addInfo,
                                        supervisor='ООКУ ТС MBH дежурный инженер',
                                        initiator='',
                                        BS=BS, describeW=describeW, performer=performer, performerG=performerG,
                                        typeNE='БС', klassificator=['14', '01', '07'],incident=incident,opercontext=opercontext,
                                        days=9)

        elif sys.argv[2] == "Интеграция":
            Remedy_Web_1_0.remedyLogin(driver)
            # for value in str1.split("\n"):
            addInfo, describeW, BS = get_param_integ(value)
            performer = ""
            performerG = 'ЮГ\ЕЦУС\УИТС\MBH Юго-Запад'
            BS = BS[4:-2]
            print(BS)
            Remedy_Web_1_0.remedy_set_param(driver, typeW='Планово-профилактическая',
                                        typeS='Данные + Голос',
                                        impactS='Нет влияния на сервис абонентам в зоне действия СЭ.',
                                        addInfo=addInfo,
                                        supervisor='ООКУ ТС MBH дежурный инженер',
                                        initiator='',
                                        BS=BS, describeW=describeW, performer=performer, performerG=performerG,
                                        typeNE='MBH', klassificator=['14', '01', '02'],incident=incident,opercontext=opercontext,
                                        days=5)
        elif sys.argv[2] == "АГОС":

            describeW = "Прошу проверить АГОС по " + dicVal['MBH']
            klassificator = ['15', '12', '10']
            performerG = "ЮГ\Краснодар\ЭТС\TN"
            days = 30
            Remedy_Web_1_0.remedyLogin(driver)
            # for value in str1.split("\n"):
            performer = ""
            if 'PL' in dicVal['PL']:
                BS=dicVal['PL'][3:]
            else:
                BS =dicVal['PL']
            Remedy_Web_1_0.remedy_set_param(driver, typeW='Планово-профилактическая',
                                        typeS='Данные + Голос',
                                        impactS='Нет влияния на сервис абонентам в зоне действия СЭ.',
                                        addInfo='',
                                        supervisor='Оперативный дежурный ООКУ СРД ФО',
                                        initiator='',
                                        BS=BS, describeW=describeW, performer=performer, performerG=performerG,
                                        typeNE='БС', klassificator=klassificator,incident=incident,opercontext=opercontext,
                                        days=28)

        elif sys.argv[2] == "Конфиг":

            describeW = "Прошу выделить int ip и VLAN  для подготовки config файла " +dicVal['Type'] +' '+ dicVal['MBH']
            # describeW = "Прошу cформировать snapshot файл для узла " +dicVal['Type'] +' '+ dicVal['MBH']+" и направить на почту yvsobol3@mts.ru"
            klassificator = ['14', '01', '01']
            performerG = 'ЮГ\ЕЦУС\УИТС\MBH Юго-Запад'
            days = 7
            Remedy_Web_1_0.remedyLogin(driver)
            # for value in str1.split("\n"):
            performer = ""
            if 'PL' in dicVal['PL']:
                BS = dicVal['PL'][3:]
            else:
                BS = dicVal['PL']
            Remedy_Web_1_0.remedy_set_param(driver, typeW='Планово-профилактическая',
                                            typeS='Данные + Голос',
                                            impactS='Нет влияния на сервис абонентам в зоне действия СЭ.',
                                            addInfo='Узел планируется подключить к ' + dicVal['To']+', тип подключения-'+dicVal['Via'] + ".Lo0:="+dicVal['lo0'] +". Информацию прошу заполнить в формате : ip:=;VLAN:=; ",
                                            supervisor='Оперативный дежурный ООКУ СРД ФО',
                                            initiator='',
                                            BS=BS, describeW=describeW, performer=performer, performerG=performerG,
                                            typeNE='БС', klassificator=klassificator,incident=incident,opercontext=opercontext,
                                            days=15)
        elif sys.argv[2] == "Чеклист":

            describeW = "Прошу проверить чек-лист по узлу " +dicVal['Type'] +' '+ dicVal['MBH']
            klassificator = ['14', '01', '06']
            performerG = 'ЮГ\ЕЦУС\УИТС\MBH Юго-Запад'
            days = 7
            Remedy_Web_1_0.remedyLogin(driver)
            # for value in str1.split("\n"):
            performer = ""
            if 'PL' in dicVal['PL']:
                BS = dicVal['PL'][3:]
            else:
                BS = dicVal['PL']
            Remedy_Web_1_0.remedy_set_param(driver, typeW='Планово-профилактическая',
                                            typeS='Данные + Голос',
                                            impactS='Нет влияния на сервис абонентам в зоне действия СЭ.',
                                            addInfo='',
                                            supervisor='Оперативный дежурный ООКУ СРД ФО',
                                            initiator='',
                                            BS=BS, describeW=describeW, performer=performer, performerG=performerG,
                                            typeNE='БС', klassificator=klassificator,incident=incident,opercontext=opercontext,
                                            days=7)

        Remedy_Web_1_0.driver_quit(driver)
        time.sleep(5)



    print("Время выполнения программы: %s секунд" % (time.time() - start_time))
    # sys.exit()