import Cedar
from selenium import webdriver
import Remedy_Web_1_0
import sys
import time
if __name__ == "__main__":
    options = webdriver.ChromeOptions()
    # options.add_argument('headless')  # если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(options=options)
    Credentials = Remedy_Web_1_0.load_obj('Cred')
    Cedar.login(driver, Credentials['login'], Credentials['pass'])
    # Получаем словарь из Excel
    for value in sys.argv[1].split("\r")[:-1]:
        list_param = value.split("-:")
        list_param.remove('')
        for n, v in enumerate(list_param):
            list_param[n] = v.split("=")
        dicVal = dict(list_param)
        print(dicVal)
        time.sleep(15)
        BSID=Cedar.getCedarBSID(driver,dicVal['PL'][3:].replace('_',''),'')
        Cedar.addMBH(driver,BSID,dicVal['MBH'],dicVal['Type'],dicVal['Date'],dicVal['IP'],dicVal['Notes'])