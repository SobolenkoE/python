import NIOSS
import Remedy_Web_1_0
import sys
import pythonping
def getIP(mail_body):
    List_strings=[]
    List_strings_proto =mail_body.split('\n')
    for string in List_strings_proto:
        Sub_string_list=string.split('\r')
        for sub_string in Sub_string_list:
            if NIOSS.stripIPadress(sub_string)!='error':
                IP=NIOSS.stripIPadress(sub_string)
                List_strings.append(IP)
    return List_strings

if __name__ == "__main__":
    try:
        performer, performerG, initiator = Remedy_Web_1_0.get_performer_by_mail(sys.argv[2],"TN")
        mail_person = performer
        List_ip = getIP(sys.argv[1])
        if performerG == 'error':
            Remedy_Web_1_0.reply_send(sys.argv[2], 'Пользователь', 'Автозаведение работ', 'Email не зарегистрирован!')
            sys.exit()
        if not List_ip:
            Remedy_Web_1_0.reply_send(sys.argv[2], 'Пользователь', 'Автозаведение работ', 'Не удалось распознать сообщение.')
            sys.exit()
    except:
        print('Ошибка распознавания в письме')
        Remedy_Web_1_0.reply_send(sys.argv[2], 'Пользователь', 'Автозаведение работ', 'Ошибка распознавания текста в письме!')
        sys.exit()
    for ip in List_ip:
        Ping=pythonping.ping(ip, size=40, count=4)
        if Ping.rtt_min_ms > 1000:
            answer=ip +' не пингуется('+'\n'
        else:
            answer = 'Пингуется!!!'+'\n'
        answer=answer+'\n'.join('{}'.format(item) for item in Ping._responses)
        Remedy_Web_1_0.reply_send(sys.argv[2],  performer, 'AutoPing', answer)