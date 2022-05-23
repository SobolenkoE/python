
import time
import Remedy_Web_1_0
import Main
import sys
# -:PL_A=PL_23_1229-:PL_Z=PL_23_0223-:describe=70/80 GHz UltraLink-FX80, 0.3/0.3, 1+0


if __name__ == "__main__":
    # sys.argv=['',"",""]
    # sys.argv[1]='-:PL_A=PL_23_0717-:PL_Z=PL_23_0465-:RRL=iPas-:describe='
    # sys.argv[2] ="Detect_load"

    opercontext = '',
    start_time = time.time()
    # options = Remedy_Web_1_0.webdriver.ChromeOptions()
    # options.add_argument('headless') #если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    incident=''
    # driver = Remedy_Web_1_0.webdriver.Chrome(options=options)
    # Credentials = Remedy_Web_1_0.load_obj('Cred')
    driver=Remedy_Web_1_0.get_driver(Remedy_Web_1_0.webdriver.ChromeOptions())
    Remedy_Web_1_0.remedyLogin(driver)

    for value in sys.argv[1].split("\r")[:-1]:
        # Получаем словарь из Excel
        list_param = value.split("-:")
        list_param.remove('')
        for n, v in enumerate(list_param):
            list_param[n] = v.split("=")
        dicVal = dict(list_param)
        # значения по умолчанию
        performer = ""
        performerG = 'ЮГ\ЕЦУС\УИТС\MBH Юго-Запад'
        addInfo=dicVal['describe']
        BS=dicVal['PL_A'][3:]
        describeW = ""
        days=3
        klassificator = ['14', '01', '07']
        # переопределяем умолчания в зависимости от аргумента 2
        if sys.argv[2]=="IP":
            describeW="Прошу выделить ip адресацию для новой РРЛ " + dicVal['PL_A'][3:] + "<>" +dicVal['PL_Z'][3:] +'. Также прошу добавить пролет в СУ и прописать управление на сети MBH.'
            klassificator = ['14', '02', '06']
            days = 14
        elif sys.argv[2]=="AGOS":
            describeW = "Прошу проверить АГОС по РРЛ " + dicVal['PL_A'][3:] + "<>" + dicVal['PL_Z'][3:]
            klassificator = ['15', '12', '10']
            performerG = "ЮГ\Краснодар\ЭТС\TN"
            days = 7
        elif sys.argv[2] == "Del":
            describeW = "В связи с приемкой РРЛ " + dicVal['PL_A'][3:] + "<>" + dicVal['PL_Z'][3:] + " прошу подготовить ранее сущ. РРЛ к демонтажу. Выписать из СУ, предоставив серийники ODU."
            klassificator = ['14', '02', '07']
            days = 7
        elif sys.argv[2] == "Switch":
            describeW = "В связи с приемкой РРЛ " + dicVal['PL_A'][3:] + "<>" + dicVal['PL_Z'][3:] + " прошу переключить нагрузку на новый пролет."
            klassificator = ['14', '02', '10']
            days = 5
        elif sys.argv[2] == "Check":
            describeW =  "Прошу проверить чек лист по пролету: " + dicVal['PL_A'][3:] + "<>" + dicVal['PL_Z'][3:] + "."
            klassificator = ['14', '02', '05']
            days = 5
        elif sys.argv[2] == "Detect_load":
            print(dicVal)
            describeW = "Прошу подготовить список нагрузки, которую необходимо перелючить для демонтажа РРЛ: " + dicVal['PL_A'][3:] + "<>" + dicVal['PL_Z'][3:] +  " тип РРЛ:" + dicVal['RRL'][3:] +"."
            klassificator = ['14', '05', '05']
            days = 5
        elif sys.argv[2] == "Detect_load_oldRRL":
            describeW = "Прошу подготовить список нагрузки, которую необходимо перелючить на новую РРЛ со старой на участке: " + dicVal['PL_A'][3:] + "<>" + dicVal['PL_Z'][3:]  +". В формате порт:-нагрузка, включая транзитный трафик."
            klassificator = ['14', '05', '05']
            days = 5
        else:
            print("Тип аргумента 2 задан неверно!")
            sys.exit()
        # заводим работу
        Remedy_Web_1_0.remedy_set_param(driver,typeW='Планово-профилактическая',
        typeS='Данные + Голос',
        impactS='Нет влияния на сервис абонентам в зоне действия СЭ.',
        addInfo=addInfo,
        supervisor='ООКУ ТС РРЛ дежурный инженер',
        initiator='',
        BS=BS, describeW=describeW,performer=performer, performerG=performerG,
                                        incident=incident,opercontext=opercontext,
        typeNE='БС',klassificator=klassificator,
                         days=days)
        time.sleep(8)
        driver.get('http://remedy.msk.mts.ru/arsys/forms/remedy-prom/I2%3AWorks/Default+Administrator+View/?cacheid=&mode=CREATE')
        try:
            alert = driver.switch_to.alert()
            alert.accept()
            driver.switch_to.default_content()
            time.sleep(8)
        except:
            print('Работа заведена без уведомлений')

    print("Время выполнения программы: %s секунд" % (time.time() - start_time))

    driver.quit()