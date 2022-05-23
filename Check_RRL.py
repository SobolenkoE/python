import pymysql
import Remedy_Web_1_0
import sys

def get_RRL_List(mail_body):
    List_strings=[]
    List_strings_proto =mail_body.split('\n')
    for string in List_strings_proto:
        Sub_string_list=string.split('\r')
        for sub_string in Sub_string_list:
            if "-" in sub_string and len(sub_string)>10 and len(sub_string)<16:
                List_strings.append(sub_string)
    return List_strings

def RRNNNNtoRR_NNNN(BS):
    BSR = ""
    for n in BS[::-1]:
        if n.isdigit():
            BSR = n + BSR
    BSR = BSR[:2] + "_" + BSR[2:]
    return BSR

def check_RRL(RRL_string):
    list_status=['none','FAIL','NOW OK','OK']
    answer=""
    RRL_string=RRL_string.replace("_","").replace("PL","")
    BS=[]
    if RRL_string.count('-')==1:
        BS=RRL_string.split('-')
        for n,BS_N in enumerate(BS):
            BS[n]=RRNNNNtoRR_NNNN(BS_N)

    con = pymysql.connect(host='10.40.254.146', user='monitor', password='monitoring123456', database='portal_tn',
                          port=3306)
    with con:
        cur = con.cursor()
        cur.execute(
            "select * from portal_tn.check_rrn where ((rrn_a like '%" +BS[0]+ "%' and rrn_b like '%" +BS[1]+ "%') OR (rrn_a like '%" +BS[1]+ "%' and rrn_b like '%" +BS[0]+ "%')) AND (region = 'Краснодарского кр.' OR region= 'Республика Адыгея')")
        Lists=cur.fetchall()
        if not Lists:
            answer = answer + 'Пролет РРЛ ' +BS[0]+ "<>"+BS[1]+'  не проверялся по чек-листу.\n'+'\n'+'\n'
        for List in Lists:
            answer = answer + "Пролет РРЛ "+ List[3]+ "<>"+List[4] + " "+List[7]
            if List[21]==1:
                answer = answer + " ПРИНЯТ по чек листу! Проверял " +List[11]
            elif List[21]==2:
                answer = answer + " НЕ ПРИНЯТ по чек листу! Проверял " + List[11]
            answer = answer + ".\nNIOSS-" + list_status[List[14]]
            answer = answer + "\nName-" + list_status[List[17]]
            answer = answer + "\nErrors-" + list_status[List[18]]
            answer = answer + "\nRSL-" + list_status[List[19]]
            answer = answer + "\nAlarm-" + list_status[List[20]]
            if List[22]:
                answer = answer + "\nКомментарий-" + List[22]
            answer= answer+'\n'+'\n'+'\n'
        return answer

if __name__ == "__main__":
    # sys.argv = ['C:\\Users\\yvsobol3\\PycharmProjects\\Test\\Check_RRL.py', '230444-230044\r \n', 'sob.evg@yandex.ru', 'checklistRRL']
    Answer = ""
    try:
        performer, performerG, initiator = Remedy_Web_1_0.get_performer_by_mail(sys.argv[2],"TN")
        mail_person = performer
        List_RRL = get_RRL_List(sys.argv[1])
        if performerG == 'error':
            Remedy_Web_1_0.reply_send(sys.argv[2], 'Пользователь', 'Автозаведение работ', 'Email не зарегистрирован!')
            sys.exit()
        if not List_RRL:
            Remedy_Web_1_0.reply_send(sys.argv[2], 'Пользователь', 'Автозаведение работ', 'Не удалось распознать сообщение.')
            sys.exit()
    except:
        print('Ошибка распознавания в письме')
        Remedy_Web_1_0.reply_send(sys.argv[2], 'Пользователь', 'Автозаведение работ', 'Ошибка распознавания текста в письме!')
        sys.exit()
    for RRL in List_RRL:
        Answer=Answer+'\n'+check_RRL(RRL)
    Remedy_Web_1_0.reply_send(sys.argv[2],  performer, 'Информация по чек-листу', Answer)