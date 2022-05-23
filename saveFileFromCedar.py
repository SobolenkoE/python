import requests
import sys
from bs4 import BeautifulSoup
import Remedy_Web_1_0
import Cedar
import pandas




def get_session(Credentials):
    try:
        session=Cedar.load_obj('session')
        response = session.get('https://cedar.mts.ru/basic/web/site', verify=False)
        if response.url != 'https://cedar.mts.ru/basic/web/site':
            raise Exception
    except:
        session = requests.session()
        response = session.get('https://cedar.mts.ru/basic/web/site/login',verify=False)
        soup = BeautifulSoup(response.text, 'lxml')
        csrf=str(soup.find_all('meta',attrs={"name": "csrf-token"})).split('="')[1].split('"')[0] # парсим токен


        payload = {
            '_csrf':csrf,
            'LoginForm[login]': Credentials['login'],
            'LoginForm[password]': Credentials['pass']
        }
        session.post('https://cedar.mts.ru/basic/web/site/login',data=payload)
    Cedar.save_obj(session,'session')
    return session

def parseArgs(args):
    try:
        reportName=args.split('отчета ')[1].split(' завершено')[0].strip()
    except:
        reportName ='Отчет без названия'
    link=args.split('ссылке ')[-1].strip()
    return link,reportName

def main(args):
    path='\\\\msk.mts.ru\\ug\\mr\\DRS\\Docs\\_ОтделРазвитияТранспортныхСетей\\Выгрузки\\Cedar\\'
    Credentials = Remedy_Web_1_0.load_obj('Cred')
    session = get_session(Credentials)
    link,report_name=parseArgs(args)
    response = session.get(link)
    with open(path+report_name+'temp.xlsx', "wb") as f:
        f.write(response.content)
    # pd_temp=pandas.read_excel(path+report_name+'temp.xlsx',header=2,index_col=0)
    # pd_temp.drop(labels = [1],inplace=True)
    # pd_old=pandas.read_excel(path+report_name+'.xlsx')
    # pd_old.drop(labels=[1], inplace=True)
    # pd=pandas.concat([pd_temp, pd_old]).drop_duplicates(keep=False)

    pass

if __name__ == "__main__":
    # Cedar.save_obj(sys.argv[1],'arvs')
    text=Cedar.load_obj('arvs')
    # main(sys.argv[1])
    main(text)