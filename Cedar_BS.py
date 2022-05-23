from selenium import webdriver

import Cedar
import Remedy_Web_1_0
import Remedy_Web
import time
import sys
import re

def dms2dd(coord):
    NPart,EPart=coord.split(', ')
    Ndirection,Ndegrees=NPart.split('°')[0].split(' ')
    Edirection, Edegrees = EPart.split('°')[0].split(' ')
    Nminutes,Nseconds=NPart.split('° ')[1].split('\' ')
    Nseconds=Nseconds.strip('"')
    Eminutes, Eseconds = EPart.split('° ')[1].split('\' ')
    Eseconds = Eseconds.strip('"')
    ddN = float(Ndegrees) + float(Nminutes)/60 + float(Nseconds)/(60*60)
    ddE = float(Edegrees) + float(Eminutes) / 60 + float(Eseconds) / (60 * 60)
    return round(ddN,6), round(ddE,6)






if __name__ == "__main__":
    text = ''
    # sys.argv='','230525\r\nС уважением\r\nВладимир Стародубцев\r\n','sob.evg@yandex.ru','getidbs'
    try:
        print(sys.argv[2])
        performer, G,initiator = Remedy_Web_1_0.get_performer_by_mail(sys.argv[2],'TN')
        List_BS = Cedar.getParam_BS(sys.argv[1])
    except:
        Cedar.reply_send(sys.argv[2], '', 'Автоматические исходные данные', 'Доступ запрещен', "")
        sys.exit()


    start_time = time.time()
    options = webdriver.ChromeOptions()
    options.add_argument('headless')  # если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(options=options)
    Credentials = Remedy_Web_1_0.load_obj('Cred')
    Cedar.login(driver, Credentials['login'], Credentials['pass'])

    for BS_A in List_BS:
        Address_A, AMS_A,Coord,Additional,Podrazdelenie = Cedar.get_data_PL(driver, BS_A, sys.argv[2])
        N,E=dms2dd(Coord)
        naviLink='<https://yandex.ru/navi/?whatshere%5Bpoint%5D='+ str(E) +'%2C'+ str(N) +'&whatshere%5Bzoom%5D=15&lang=ru&from=navi>'
        text =text+  ("\nПлощадка " + BS_A )
        text =text+ ("\nАдрес: %s" % Address_A)
        text =text+ ("\nТип АМС: %s" %AMS_A)
        text =text+ ("\nКоординаты в WGS: %s" % Coord)
        text =text+ ("\nБрать ключи в: %s" % Podrazdelenie)
        text =text+ ("\nИнформация по проходу: %s" % Additional)
        text =text + ("\nПроехать %s" % naviLink)
    text = "Запрос отработан успешно!" + text
    Cedar.reply_send(sys.argv[2], performer, 'Автоматические исходные данные', text,'')
    driver.quit()
    sys.exit()
