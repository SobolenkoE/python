import re, os, time, datetime, uuid
import Remedy_Web_sr

if __name__ == "__main__":
    while True:
        try:
            doc_file = open('C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user\\Work_list.txt', 'r+')
            file_data = doc_file.read()
            doc_file.close()
            doc_file = open('C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user\\Work_list.txt', 'w+')
            doc_file.close()
            works_from_file=file_data.split('{')[1:]
            if works_from_file:
                driver = Remedy_Web_sr.get_driver(Remedy_Web_sr.webdriver.ChromeOptions())
                Remedy_Web_sr.remedyLogin(driver)
                for work in works_from_file:
                    arg1=work.split("email:=")[0].split("body:=")[1]
                    work=work.split("email:=")[1]
                    arg2 = work.split("item:=")[0].split("\n")[0]
                    arg3 = work.split("item:=")[1].split("\n")[0]
                    try:
                        Remedy_Web_sr.work_create(driver,arg1,arg2,arg3)
                    except:
                        driver.quit()
                        driver = Remedy_Web_sr.get_driver(Remedy_Web_sr.webdriver.ChromeOptions())
                        Remedy_Web_sr.remedyLogin(driver)
                        # записываем в файл работу, которую не удалось завести для обработки в следующем цикле
                        doc_file = open('C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user\\Work_list.txt', 'r+')
                        file_data = doc_file.read()
                        doc_file.close()
                        doc_file = open('C:\\Users\\yvsobol3\\PycharmProjects\\Test\\user\\Work_list.txt', 'w+')
                        doc_file.write(file_data+"\n{"+work)
                        doc_file.close()
                Remedy_Web_sr.driver_quit(driver)
        except:
            # try:
            #     # driver.quit()
            # except:
                pass

        time.sleep(10)
