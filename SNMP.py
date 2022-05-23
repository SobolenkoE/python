from pysnmp.hlapi import *
from pysnmp.entity.rfc3413.oneliner import cmdgen
NumberIntToShort = {"1":"41","2":"42","3":"43","4":"44","5":"45","6":"46","7":"47","8":"48","9":"51","10":"52","11":"53","12":"54"}
NumberIntToLong = {"1":"16842752","2":"25231360","3":"33619968","4":"42008576","5":"50397184","6":"58785792","7":"67174400","8":"75563008","9":"100728832","10":"109117440","11":"117506048","12":"125894656"}
NumberIntToModem = {"1": "1", "2": "2", "3": "3", "4": "4", "5": "5", "6": "6", "7": "7", "8": "8",
                                 "9": "11", "10": "12", "11": "13", "12": "14"}
HuaweiCodeToType={'Huawei Optix RTN V100R019C10SPC200':'Huawei RTN 380A'}


# function section

def errors(errorIndication, errorStatus):
    #  обработка ошибок в случае ошибок возвращаем False и пишем в файл file
    if errorIndication:
        print(errorIndication)
        return True
    elif errorStatus:
        errorStatus.prettyPrint()
        return True
    else:
        return False

# функция для запроса значения по OID
def snmp_getcmd(community, ip, port, OID):
    errorIndication, errorStatus, errorIndex, varBinds = next(getCmd(SnmpEngine(),
                   CommunityData(community),
                   UdpTransportTarget((ip, port)),
                   ContextData(),
                   ObjectType(ObjectIdentity(OID))))
    # print(OID, errorIndication, errorStatus)
    if not errors(errorIndication, errorStatus):
        for name,val in varBinds:
            return (val.prettyPrint())
    else:
            return ('Error')

# функция для запроса значения по OID
def snmp_walk(community, ip, port, OIDs, OIDp):
    OID = OIDs.split(".")
    OID = OID[:-1]
    OID = ".".join(OID)
    result = []
    cmdGen = cmdgen.CommandGenerator()

    errorIndication, errorStatus, errorIndex, varBindTable = cmdGen.nextCmd(
        cmdgen.CommunityData(community),
        cmdgen.UdpTransportTarget((ip, port)),
        OID
    )

    if errorIndication:
        print(errorIndication)
    else:
        if errorStatus:
            return ('%s at %s' % (
                errorStatus.prettyPrint(),
                errorIndex and varBindTable[-1][int(errorIndex) - 1] or '?'
            )
                    )
        else:
            for varBindTableRow in varBindTable:
                for name, val in varBindTableRow:
                    if OIDs < "." + str(name) < OIDp:
                        result.append([name, val])
            return result


# функция определяющая комьюнити по ip адерсу
def get_community(ip):
    try:
        # для iPas
        if snmp_getcmd('KCMonThy0x56', ip, 161, '.1.3.6.1.4.1.119.2.3.69.1.1.13.0')!='Error':
            return 'KCMonThy0x56'
        elif snmp_getcmd('public', ip, 161, '.1.3.6.1.4.1.119.2.3.69.1.1.13.0')!='Error':
            return 'public'
        # для MiniLink
        elif snmp_getcmd('KCMonThy0x56', ip, 161, '.1.3.6.1.2.1.47.1.1.1.1.2.1') != 'Error':
            return 'KCMonThy0x56'
        elif snmp_getcmd('public', ip, 161, '.1.3.6.1.2.1.47.1.1.1.1.2.1') != 'Error':
            return 'public'

        # для Intracom и Huawei
        elif snmp_getcmd('KCMonThy0x56', ip, 161, '.1.3.6.1.2.1.1.1.0') != 'Error':
            return 'KCMonThy0x56'
        elif snmp_getcmd('public', ip, 161, '.1.3.6.1.2.1.1.1.0') != 'Error':
            return 'public'
        else:
            return 'none'

    except Exception as e:
        print('Оборудование по указанному ip недоступно')

# функция определяющая комьюнити по ip адерсу
def get_typeRRL(ip,community):
    dictypeipas={'400':'400','210':'200','1000':'1000','520':'EX','100':'NEO ST','10520':'EX adv','170':'NEO HP'}
    typeRRL=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.1.1.13.0') #для iPas
    if typeRRL!='Error' and typeRRL in dictypeipas:
        return 'iPasolink ' + dictypeipas[typeRRL]
    else:
        typeRRL = snmp_getcmd(community, ip, 161, '.1.3.6.1.2.1.1.1.0') #для ML TN & OmniBas
        if typeRRL != 'Error':
            if typeRRL in ['MINI-LINK Traffic Node','MINI-LINK 6600']:
                typeRRL = snmp_getcmd(community, ip, 161, '.1.3.6.1.2.1.47.1.1.1.1.2.1') #для ML TN
            return typeRRL
        else:
            return 'none'
# для CX другие OIDы

def GetFreqTXOBCX(ip,community):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.2.3.1.3.1')

    if freqTX!='Error':
        return freqTX
    else:
        return 'none'

def GetFreqRXOBCX(ip,community):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.2.3.1.8.1')
    if freqTX!='Error':
        return freqTX
    else:
        return 'none'


# функция по запросу имени
def GetNameUL(ip,community):
    TX=snmp_getcmd(community, ip, 161, '.1.3.6.1.2.1.1.5.0')
    if TX!='Error':
        return TX
    else:
        return 'none'

# функция по запросу частоты у ультралинка
def GetFreqTXUL(ip,community):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.2.3.1.3.1')

    if freqTX!='Error':
        return freqTX
    else:
        return 'none'

# функция по запросу частоты у ультралинка
def GetFreqRXUL(ip,community):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.2.3.1.8.1')
    if freqTX!='Error':
        return freqTX
    else:
        return 'none'

# функция по запросу мощности передатчика у ультралинка
def GetPowerTXUL(ip,community):
    TX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.2.4.1.5.1')
    if TX!='Error':
        return str(int(TX)/10)
    else:
        return 'none'

# функция по запросу мощности приемника у ультралинка
def GetPowerRXUL(ip,community):
    TX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.2.4.1.3.1')
    if TX!='Error':
        return "-" + str(int(TX)/10)
    else:
        return 'none'

# функция по запросу ширины спектра у ультралинка
def GetCannel_SpacingUL(ip, community):
    dic={"0":"7","1":"14","2":"28","3":"56","4":"40","5":"20","7":"250","8":"500","14":"1000"}
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.2.1.1.2.1')
    if TX != 'Error':
        return dic[TX]
    else:
        return 'none'

# функция по запросу айпишника ответной стороны у ультралинка
def GetOppositeIPUL(ip, community):
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.1.9.1.3.1')
    if TX != 'Error':
        return TX
    else:
        return 'none'

# функция по запросу HiLow у ультралинка
def GetUpperUL(ip, community):
    dic = {"1": "Lower", "0": "Upper"}
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.2.2.1.2.1')
    if TX != 'Error':
        return dic[TX]
    else:
        return 'none'

# функция по запросу Модуляции и емкости у ультралинка
def GetModulationUL(ip, community):
    dicMod = {"4QAM": "4 QAM", "4QAMhigh": "4 QAM","16QAM": "16 QAM", "32QAM": "32 QAM","64QAM": "64 QAM", "128QAM": "128 QAM","256QAM": "256 QAM","512QAM": "512 QAM", "1024QAM": "1024 QAM","2048QAM": "2048 QAM"}
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.1.3.1.7.1')
    TXsecond=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.1.2.1.2.' + TX)
    if TXsecond != 'Error':
        return dicMod[TXsecond.split(" ")[1]],TXsecond.split(" ")[4]
    else:
        return 'none'

#Функция по
def GetPolariztionUL(ip, community):
    dicXPIC = {"1": "1+0", "2": "XPIC"}
    dicPol = {"1": "Single", "2": "X-Поляризация"}
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.1807.55.1.1.14.1.2.1')
    if TX!= 'Error':
        try:
            return dicPol[TX],dicXPIC[TX]
        except:
            return "1+0", "Single"
    else:
        return 'none'
#code section

# функция по запросу имени
def GetNameiPas(ip,community):
    TX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.5.1.1.1.3.1')
    if TX!='Error':
        return TX
    else:
        return 'none'

# //////////////  NEO
# //////////////  NEO
# //////////////  NEO
def GetNameiPasNEO(ip,community):
    TX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.3.1.1.1.0')
    if TX!='Error':
        return TX
    else:
        return 'none'

def GetOppositeIPiPasNEO(ip, community):
    TX = str(snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.1.1.9.0'))
    if TX != 'Error' and TX:
        return TX
    else:
        return 'none'


# фунция по получению частоты iPasNEO HP
def GetFreqTXiPasNEOHP(ip,community):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.407.4.2.1.2.1')

    if freqTX!='Error':
        return str(float(freqTX))
    else:
        return 'none'

def GetFreqRXiPasNEOHP(ip,community):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.407.4.2.1.3.1')
    if freqTX!='Error':
        return str(float(freqTX))
    else:
        return 'none'



# фунция по получению частоты iPasNEO
def GetFreqTXiPasNEO(ip,community):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.401.4.2.1.2.1')

    if freqTX!='Error':
        return str(float(freqTX))
    else:
        return 'none'

# фунция по получению частоты iPas
def GetFreqRXiPasNEO(ip,community):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.401.4.2.1.3.1')
    if freqTX!='Error':
        return str(float(freqTX))
    else:
        return 'none'

def GetPowerTXiPasNEOHP(ip,community):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.407.8.1.1.3.1')
    if freqTX!='Error':
        return freqTX
    else:
        return 'none'

    # фунция по получению мощности iPas
def GetPowerRXiPasNEOHP(ip, community):
        freqTX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.407.8.1.1.5.1')
        if freqTX != 'Error':
            return freqTX
        else:
            return 'none'



# фунция по получению мощности iPas
def GetPowerTXiPasNEO(ip,community):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.401.8.1.1.3.1')
    if freqTX!='Error':
        return freqTX
    else:
        return 'none'

    # фунция по получению мощности iPas
def GetPowerRXiPasNEO(ip, community):
        freqTX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.401.8.1.1.5.1')
        if freqTX != 'Error':
            return freqTX
        else:
            return 'none'

def GetCannel_SpacingiPasNEO(ip, community):
    dic={"1":"3.5","2":"7","3":"14","4":"28","5":"56","9":"250","10":"500"}
    TX = "4"
    if TX != 'Error':
        return dic[TX]
    else:
        return 'none'

def GetUpperiPasNEO(ip, community):
    dic={"1":"Lower","2":"Upper"}
    TX=GetFreqTXiPasNEO(ip, community)
    RX=GetFreqRXiPasNEO(ip, community)
    if RX<TX:
        return dic["2"]
    else:
        return dic["1"]


def GetSubbandPasNEO(ip, community):
    TX = ""
    if TX != 'Error':
        return TX
    else:
        return 'none'

def GetPolarizationiPasNEO(ip, community):
    dicXPIC = {"No Such Instance currently exists at this OID": "1+0", "2": "1+1", "4": "XPIC"}
    dicPol = {"No Such Instance currently exists at this OID": "Single", "4": "X-Поляризация","2":"Single"}
    TX = 'No Such Instance currently exists at this OID'
    if TX != 'Error':
        return dicPol[TX],dicXPIC[TX]
    else:
        return 'none'

def GetCapacityiPasNEO(ip, community):
    TX = ""
    if TX != 'Error':
        return TX
    else:
        return 'none'

def GetModulationiPasNEO(ip, community):
    dicMod = {"10": "4 QAM", "11": "16 QAM", "12": "32 QAM", "13": "64 QAM",
              "14": "128 QAM", "15": "256 QAM", "33": "512 QAM", "34": "1024 QAM"}
    TX=[]
    if TX != 'Error':
        return dicMod["11"], dicMod["12"]
    else:
        return 'none'









# функция по запросу айпишника ответной стороны у iPasADV
def GetOppositeIPiPasADV(ip, community,modem):

    TX = str(snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.5.4.1.1.3.'+NumberIntToShort[modem]))
    if TX != 'Error' and TX:
        TX ='.'.join(map(lambda a: str(int(a,16)),[TX[6:8],TX[8:10],TX[10:12],TX[12:14]]))
        return TX
    else:
        return 'none'

def GetCapacityiPasADV(ip, community, modem):
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.5.21.22.1.3.'+NumberIntToLong[modem])
    if TX != 'Error':
        return str(int(TX)/1000)
    else:
        return 'none'



# //////////////// iPas
# //////////////// iPas
# //////////////// iPas
# функция по запросу айпишника ответной стороны у iPas
def GetOppositeIPiPas(ip, community,modem):
    TX = str(snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.5.4.1.1.3.'+NumberIntToShort[modem]))
    if TX != 'Error' and TX:
        TX= str(int(TX[2:4],16)) + '.' + str(int(TX[4:6], 16))+ '.' + str(int(TX[6:8], 16))+ '.' + str(int(TX[8:10], 16))
        return TX
    else:
        return 'none'




# фунция по получению частоты iPas
def GetFreqTXiPas(ip,community,modem):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.4.2.1.3.'+NumberIntToLong[modem])
    if freqTX!='Error':
        return str(float(freqTX))
    else:
        return 'none'

# фунция по получению частоты iPas
def GetFreqRXiPas(ip,community,modem):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.4.2.1.4.'+NumberIntToLong[modem])
    if freqTX!='Error':
        return str(float(freqTX))
    else:
        return 'none'


# фунция по получению мощности iPas
def GetPowerTXiPas(ip,community,modem):
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.8.1.1.4.'+NumberIntToLong[modem])
    if freqTX!='Error':
        return freqTX
    else:
        return 'none'

    # фунция по получению мощности iPas
def GetPowerRXiPas(ip, community, modem):
        freqTX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.8.1.1.6.' + NumberIntToLong[modem])
        if freqTX != 'Error':
            return freqTX
        else:
            return 'none'

def GetCannel_SpacingiPas(ip, community, modem):
    dic={"1":"3.5","2":"7","3":"14","4":"28","5":"56","9":"250","10":"500"}
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.4.2.1.8.'+NumberIntToLong[modem])
    if TX != 'Error':
        return dic[TX]
    else:
        return 'none'

def GetUpperiPas(ip, community, modem):
    dic={"1":"Lower","2":"Upper"}
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.7.3.1.25.'+NumberIntToLong[modem])
    if TX != 'Error':
        return dic[TX]
    else:
        return 'none'

def GetSubbandPas(ip, community, modem):
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.7.3.1.28.'+NumberIntToLong[modem])
    if TX != 'Error':
        return TX
    else:
        return 'none'

def GetPolarizationiPas(ip, community, modem):
    dicXPIC = {'No Such Object currently exists at this OID':'1+0',"No Such Instance currently exists at this OID": "1+0", "2": "1+1", "4": "XPIC"}
    dicPol = {'No Such Object currently exists at this OID':'Single',"No Such Instance currently exists at this OID": "Single", "4": "X-Поляризация","2":"Single"}
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.4.3.1.4.'+NumberIntToLong[modem])
    if TX != 'Error':
        return dicPol[TX],dicXPIC[TX]
    else:
        return 'none'



def GetCapacityiPas(ip, community, modem):
    TX = snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.5.21.5.1.13.'+NumberIntToLong[modem])
    if TX != 'Error' and "No Such" not in TX:
        return str(int(TX)/1000)
    else:
        return '0'

def GetModulationiPas(ip, community, modem):
    dicMod = {"10": "4 QAM", "11": "16 QAM", "12": "32 QAM", "13": "64 QAM",
              "14": "128 QAM", "15": "256 QAM", "33": "512 QAM", "34": "1024 QAM"}
    TX=[]
    for i in dicMod:
        TXtemp=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.119.2.3.69.501.4.2.1.' +i +'.'+ NumberIntToLong[modem])
        if TXtemp=='2':
            TX.append([i,TXtemp])

    if TX != 'Error':
        return dicMod[min(TX, key=lambda item: item[0])[0]], dicMod[max(TX, key=lambda item: item[0])[0]]
    else:
        return 'none'

# функция по запросу частоты у ML TN
def GetFreqTXML(ip,community,interface):
    interface=str(int(interface)-8192)
    freqTX=str((int(snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.3.4.3.1.2.1.1.'+interface)) + int(snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.3.4.3.1.2.1.3.'+interface))*int(snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.3.4.3.1.2.1.6.'+interface)))/1000)
    if freqTX!='Error':
        return freqTX
    else:
        return 'none'

# функция по запросу частоты у ML TN
def GetFreqRXML(ip,community,interface):
    interface = str(int(interface) - 8192)
    freqTX = str((int(snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.3.4.3.1.2.1.2.' + interface)) + int(
        snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.3.4.3.1.2.1.3.' + interface)) * int(
        snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.3.4.3.1.2.1.6.' + interface))) / 1000)
    if freqTX != 'Error':
        return freqTX
    else:
        return 'none'

# фунция по получению мощности ML
def GetPowerTXML(ip,community,interface):
    interface = str(int(interface) - 8192)
    freqTX=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.3.4.3.1.3.1.1.'+interface)
    if freqTX!='Error':
        return freqTX
    else:
        return 'none'

    # фунция по получению мощности ML
def GetPowerRXML(ip, community, interface):
        interface = str(int(interface) - 8192)
        freqTX = str(int(snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.3.4.3.1.3.1.10.' + interface))/10)
        if freqTX != 'Error':
            return freqTX
        else:
            return 'none'


def GetCannel_SpacingML(ip, community, interface):
    # interface = str(int(interface) - 8192)
    dic = {"2": "14", "4": "28",  "3": "56","1": "28"}
    freqTX = snmp_getcmd(community, ip, 161, '.1.3.6.1.2.1.31.1.1.1.15.' + interface)
    if freqTX != 'Error':
        return dic[freqTX[0]]
    else:
        return 'none'

def GetUpper(self):
    if self.RRS.FreqTX>self.RRS.FreqRX:
        return "Upper"
    else:
        return "Lower"

# функция по считываению Subband
# OID=.1.3.6.1.2.1.47.1.3.2.1.2.1963204385.0, Type=OID, Value=1.3.6.1.2.1.2.2.1.1.2134638722
# на основании walk составляем таблицу соответсвия значения интерфейса в OID и в Value

def GetSubbandML(ip,community,interface):
    interface = str(int(interface) - 8192)
    interface_for_search=""
    OID_Val_Table = snmp_walk(community,ip, 161, ".1.3.6.1.2.1.47.1.3.2.1.2.1900000000",
                                   ".1.3.6.1.2.1.47.1.3.2.1.2.2000000000")
    for OID, Val in OID_Val_Table:
        OID_SPLIT = str(OID).split(".")
        Val_Split=str(Val).split(".")
        if interface==Val_Split[-1]:
            interface_for_search=snmp_getcmd(community, ip, 161, '.1.3.6.1.2.1.47.1.1.1.1.4.' + OID_SPLIT[-2])
            break
    try:
        return snmp_getcmd(community, ip, 161, '.1.3.6.1.2.1.47.1.1.1.1.2.' + interface_for_search)
    except:
        return ""


def GetPolarizationML(ip, community, modem):
    dicXPIC = {'1': "1+0", "2": "1+1", "4": "XPIC"}
    dicPol = {"1": "Single", "4": "X-Поляризация","2":"Single"}
    TX = "1" # надо написать фунцию по выявлению ЛАГ и по выявлению XPIC
    if TX != 'Error':
        return dicPol[TX],dicXPIC[TX]
    else:
        return 'none'

def GetCapacityML(ip,community,interface):
    interface = str(int(interface) - 8192)
    min_capacity=""
    max_capacity = ""
    interface_for_search=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.2.6.8.3.1.1.' + interface)
    interface_for_search =str(interface_for_search).split('.')[-1]
    if interface_for_search!='No Such Instance currently exists at this OID':
        try:
            min_capacity=str(float(snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.4.1.1.1.13.1.6.' + interface_for_search))/1000)
            max_capacity = str(
                float(snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.4.1.1.1.13.1.5.' + interface_for_search)) / 1000)
            return min_capacity,max_capacity
        except:
            return "",""
    else:
            return "", ""

def GetModulationML(ip,community,interface):
    dicMod = {"0": "4 QAM", "2": "16 QAM","3": "32 QAM", "4": "64 QAM",
              "5": "128 QAM", "6": "256 QAM", "7": "512 QAM", "1": "1024 QAM"}
    modulation=snmp_getcmd(community, ip, 161, '.1.3.6.1.4.1.193.81.3.4.1.1.16.1.5.' + interface)
    Mod_List=[]
    if modulation[:2] != 'No':
        modulation = bin(int(modulation[2:4], 16))
        for i,n in enumerate(modulation[2:]):
            if n=='1':
                Mod_List.append(str(i))
        try:
            return dicMod[min(Mod_List)],dicMod[max(Mod_List)]
        except:
            return "",""
    else:
        return "", ""

if __name__ == "__main__":
    ip_address_host = '10.42.182.201'  # From file
    community_string = get_community(ip_address_host)  # From file
    port_snmp = 161
    s1=get_typeRRL(ip_address_host,community_string)
    OID_sysName = '.1.3.6.1.4.1.119.2.3.69.1.1.13.0'  # From SNMPv2-MIB hostname/sysname
    sysname = (snmp_getcmd(community_string, ip_address_host, port_snmp, OID_sysName))
    print('hostname= ' + sysname)

