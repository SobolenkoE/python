from selenium import webdriver

import Cedar
import Remedy_Web_1_0
import Remedy_Web
import time
import sys
def enable_download_headless(browser,download_dir):
    browser.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd':'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_dir}}
    browser.execute("send_command", params)




if __name__ == "__main__":
    text = ''
    sys.argv='','href','report','path'
    href=sys.argv[1]
    print(sys.argv[1])
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": "c://test//",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False
    })
    # options.add_argument("download.default_directory=c://test//")
    # options.add_argument('headless')  # если хотим запустить chrome недивимкой
    options.add_argument("window-size=1920x1080")
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(options=options)
    Credentials = Remedy_Web_1_0.load_obj('Cred')
    enable_download_headless(driver, "c://test//")
    Cedar.login(driver, Credentials['login'], Credentials['pass'])
    driver.get("https://cedar.mts.ru/basic/web/reports/report/download?id=200730")

