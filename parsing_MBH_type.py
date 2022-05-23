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

from selenium import webdriver

def get_H(h):
    try:
        h=h.split('^')[0]
    except:
        h=str(h)
    return h

def set_prioritet(driver,object_id,value):
    nioss_el = driver.find_element_by_id('id__7_9040876642013351551_'+object_id)
    if not nioss_el.get_attribute('value'):
        NIOSS.sel_value(driver,'id__7_9040876642013351551_'+object_id,value)

    nioss_el = driver.find_element_by_id('id__7_9131576724813785131_' + object_id)
    if not nioss_el.get_attribute('value'):
        NIOSS.sel_value(driver, 'id__7_9131576724813785131_' + object_id, 'Да')


if __name__ == "__main__":
    rb = xlrd.open_workbook('C:/Users/yvsobol3/PycharmProjects/Test/upload.xlsx')
    sheet = rb.sheet_by_index(0)
    start_time = time.time()

    options = webdriver.ChromeOptions()
    options.add_argument('headless') #если хотим запустить chrome недивимкой
    # options.add_argument("window-size=1920x1080")
    driver = webdriver.Chrome(options=options)
    Credentials = Remedy_Web_1_0.load_obj('Cred')

    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        if row[0]:
            ID1=row[0]
            # ID2 = row[2]
            H1=get_H(row[1])
            D_A =row[2]

            try:
                NIOSS.NiossOpenRRL(driver,row[0])
                # driver.find_element_by_id('pcEdit').click()
                set_prioritet(driver,ID1,'5')
                NIOSS.sel_value(driver, 'id__7_9133814060013745752_' + ID1, D_A)
                NIOSS.set_value(driver, 'id__3_9133816070113745934_' + ID1, H1)

                # NIOSS.NiossOpenRRL(driver, row[2])
                # # driver.find_element_by_id('pcEdit').click()
                # set_prioritet(driver, ID2, '5')
                # NIOSS.set_value(driver, 'id__3_9133816070113745934_' + ID2, H2)
                print('строка '+str(rownum)+' OK')
            except:
                print('строка ' + str(rownum) + ' FAIL')
