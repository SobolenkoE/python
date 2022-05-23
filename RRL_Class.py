import SNMP
import NIOSS
class RRS():
    def __init__(self):
        self.Name = ""
        self.NameAsPL = ""
        self.IfIndex = ""
        self.Channel_Spasing = ""
        self.IsModemUsed=False
        self.GRP=""
        self.Capacity_min=""
        self.Capacity_max=""
        self.OduType=""
        self.AmrMin=""
        self.AmrMax=""
        self.OppositeIp=''
        self.OppositeName=''
        self.ModemInGroup=''
        self.FreqTX=""
        self.FreqRX = ""
        self.PowerTX = ""
        self.PowerRX = ""
        self.Band = ""
        self.SubBand=""
        self.Polarization = ""
        self.Reserv = ""
        self.Upper = ""

class RRL:
    def __init__(self):
        self.TypeIDU=""
        self.Vendor=""
        self.Ip = ""
        self.Name = ""
        self.Snmp_Community = "public"
        self.RRS = RRS()
        self.DictTypeIDUtoCountModem = {"OmniBAS-8W": "8", "OmniBAS-4W": "4", "OmniBAS-4W v2": "4", "OmniBAS-4Wv2": "4", "OmniBAS-2W": "2","OMNIBAS-2WCX": "2","OmniBAS-2W 16E1": "2", "ULINK-FX80-V2": "1",
                                        "iPasolink 1000": "12", "iPasolink 400": "4", "iPasolink 200": "2",
                                        "iPasolink EX": "1",'STRN-PTP-6250':'1',"iPasolink EX adv": "1","iPasolink NEO": "1","iPasolink NEO HP": "1"}
    def getTypeIDU(self):
        self.TypeIDU=SNMP.get_typeRRL(self.Ip,self.Snmp_Community)

    def GetBandFromFreq(self):
        print(self.RRS.FreqRX)
        if 7000<float(self.RRS.FreqRX)<7800: return '7'
        elif 7900<float(self.RRS.FreqRX)<8500: return '8'
        elif 10000 < float(self.RRS.FreqRX) < 12000: return '11'
        elif 17000 < float(self.RRS.FreqRX) < 19900:
            return '18'
        elif 70000 < float(self.RRS.FreqRX) < 90000: return '80'

    def getModemIndexOB(self,RRL_NIOSS):
        listNameOpposite = []
        listIpOpposite = []
        ListModemInDirection=[]
        SubList=[]
        count=int(self.DictTypeIDUtoCountModem[self.TypeIDU])+1
        for i in range(1,count):
            listIpOpposite.append(SNMP.GetOppositeIPUL(self.Ip,self.Snmp_Community+"."+str(i)))
            try:
                listNameOpposite.append(SNMP.GetNameUL(listIpOpposite[i-1],self.Snmp_Community))
                if NIOSS.RRNtoPL(listNameOpposite[i-1]) in (RRL_NIOSS.PL_A,RRL_NIOSS.PL_B):
                    SubList=[i,NIOSS.RRNtoPL(listNameOpposite[i-1])]
                    ListModemInDirection.append(SubList)
            except:
                listNameOpposite.append('none')

        if len(ListModemInDirection)==0:
            print('Не нашлось модема в нужном направлении, проверьте \nкорректность наименования IDU и соответсвие тракта в НИОСС')
        elif len(ListModemInDirection)==1:
            return str(ListModemInDirection[0][0])
        elif len(ListModemInDirection) >1:
            strM='В указанном направлении работает несколько направлений:\n'
            for SubList in ListModemInDirection:
                strM = strM+'Модем номер ' +str(SubList[0]) +' в направлении ' +SubList[1]+'\n'
            strM = strM + 'Укажите номер модема:'+'\n'
            return input(strM)

    def getModemIndexiPas(self,RRL_NIOSS):
        listNameOpposite = []
        listIpOpposite = []
        ListModemInDirection=[]
        SubList=[]
        count=int(self.DictTypeIDUtoCountModem[self.TypeIDU])+1
        for i in range(1,count):
            listIpOpposite.append(SNMP.GetOppositeIPiPas(self.Ip,self.Snmp_Community,str(i)))
            try:
                if listIpOpposite[i-1]:
                    listNameOpposite.append(SNMP.GetNameiPas(listIpOpposite[i-1],self.Snmp_Community))
                    if NIOSS.RRNtoPL(listNameOpposite[i-1]) in (RRL_NIOSS.PL_A,RRL_NIOSS.PL_B):
                        SubList=[i,NIOSS.RRNtoPL(listNameOpposite[i-1])]
                        ListModemInDirection.append(SubList)
            except:
                listNameOpposite.append('none')

        if len(ListModemInDirection)==0:
            print('Не нашлось модема в нужном направлении, проверьте \nкорректность наименования IDU и соответсвие тракта в НИОСС')
        elif len(ListModemInDirection)==1:
            return str(ListModemInDirection[0][0])
        elif len(ListModemInDirection) >1:
            strM='В указанном направлении работает несколько направлений:\n'
            for SubList in ListModemInDirection:
                strM = strM+'Модем номер ' +str(SubList[0]) +' в направлении ' +SubList[1]+'\n'
            strM = strM + 'Укажите номер модема:'+'\n'
            return input(strM)


    def getModemIndexiPasADV(self,RRL_NIOSS):
        listNameOpposite = []
        listIpOpposite = []
        ListModemInDirection=[]
        SubList=[]
        count=int(self.DictTypeIDUtoCountModem[self.TypeIDU])+1
        for i in range(1,count):
            listIpOpposite.append(SNMP.GetOppositeIPiPasADV(self.Ip,self.Snmp_Community,str(i)))
            try:
                if listIpOpposite[i-1]:
                    listNameOpposite.append(SNMP.GetNameiPas(listIpOpposite[i-1],self.Snmp_Community))
                    if NIOSS.RRNtoPL(listNameOpposite[i-1]) in (RRL_NIOSS.PL_A,RRL_NIOSS.PL_B):
                        SubList=[i,NIOSS.RRNtoPL(listNameOpposite[i-1])]
                        ListModemInDirection.append(SubList)
            except:
                listNameOpposite.append('none')

        if len(ListModemInDirection)==0:
            print('Не нашлось модема в нужном направлении, проверьте \nкорректность наименования IDU и соответсвие тракта в НИОСС')
        elif len(ListModemInDirection)==1:
            return str(ListModemInDirection[0][0])
        elif len(ListModemInDirection) >1:
            strM='В указанном направлении работает несколько направлений:\n'
            for SubList in ListModemInDirection:
                strM = strM+'Модем номер ' +str(SubList[0]) +' в направлении ' +SubList[1]+'\n'
            strM = strM + 'Укажите номер модема:'+'\n'
            return input(strM)

    # c Минилинком все сложно несколько связок определяются из разных таблиц, получаемых по snmp_walk
    def getModemIndexML(self,RRL_NIOSS):

        listOfInterfases = [] #лист с интерфейсами, где в каждой строке содержатся в виде листа номера WAN,LAG,IF,IPopposite
        listNameOpposite = []
        listIpOpposite = []
        ListModemInDirection=[]
        SubList=[]
        # заполняем  WAN & LAG номера интерфейсов
        OID_Val_Table=SNMP.snmp_walk(self.Snmp_Community,self.Ip,161,".1.3.6.1.2.1.77.1.1.1.1.2147410000",".1.3.6.1.2.1.77.1.1.1.1.2147420000")
        for OID,Val in OID_Val_Table:
            OID_SPLIT=str(OID).split(".")
            listOfInterfases.append([OID_SPLIT[-1],OID_SPLIT[-2]])

        # заполняем по WAN интерфейсам ip
        OID_Val_Table = SNMP.snmp_walk(self.Snmp_Community, self.Ip, 161, ".1.3.6.1.2.1.31.1.1.1.1.2147425000",
                                       ".1.3.6.1.2.1.31.1.1.1.1.2147428000")
        for OID,Val in OID_Val_Table:
            OID_SPLIT=str(OID).split(".")
            for item,int_wan in enumerate(listOfInterfases):
                if int_wan[0]==OID_SPLIT[-1]:
                    Name_OPP=SNMP.GetNameUL(str(Val).split("-")[-1],self.Snmp_Community)
                    listOfInterfases[item]=[*listOfInterfases[item][0:2],str(Val).split("-")[-1],Name_OPP]



        # дополняем номерами интерфейсов, которые получаем, через промежуточные номера
        OID_Val_Table = SNMP.snmp_walk(self.Snmp_Community, self.Ip, 161, ".1.3.6.1.2.1.31.1.2.1.3.2147410000",
                                       ".1.3.6.1.2.1.31.1.2.1.3.2147420000")
        for OID, Val in OID_Val_Table:
            OID_SPLIT = str(OID).split(".")

            for item, int_wan in enumerate(listOfInterfases):
                if int_wan[1] == OID_SPLIT[-2]:
                    listOfInterfases[item] = [*listOfInterfases[item][0:4],
                                              OID_SPLIT[-1]]


        for i in range(0,len(listOfInterfases)):
            try:
                if NIOSS.RRNtoPL(listOfInterfases[i][3]) in (RRL_NIOSS.PL_A,RRL_NIOSS.PL_B):
                    ListModemInDirection.append(listOfInterfases[i])
            except:
                listNameOpposite.append('none')

        if len(ListModemInDirection)==0:
            print('Не нашлось модема в нужном направлении, проверьте \nкорректность наименования IDU и соответсвие тракта в НИОСС')
        elif len(ListModemInDirection)==1:
            return ListModemInDirection[0]
        elif len(ListModemInDirection) >1:
            strM='В указанном направлении работает несколько направлений:\n'
            for item,SubList in enumerate(ListModemInDirection):
                strM = strM+'Модем номер ' +str(item) +' в направлении ' +SubList[3]+'\n'
            strM = strM + 'Укажите номер модема:'+'\n'
            return ListModemInDirection[int(input(strM))]


    def getRrsProperties(self,RRL_NIOSS):
        if self.TypeIDU in ["ULINK-FX80-V2",'UL-GX80','STRN-PTP-6250','UL-FX80-V2','UL-FX80-CN']:
            self.RRS.FreqTX=SNMP.GetFreqTXUL(self.Ip,self.Snmp_Community)
            self.RRS.FreqRX=SNMP.GetFreqRXUL(self.Ip,self.Snmp_Community)
            if self.TypeIDU in ['UL-FX80-CN']:
                self.RRS.FreqTX=NIOSS.StripFreqOB(int(self.RRS.FreqTX) / 1000)
                self.RRS.FreqRX = NIOSS.StripFreqOB(int(self.RRS.FreqRX) / 1000)
            self.RRS.PowerTX=SNMP.GetPowerTXUL(self.Ip,self.Snmp_Community)
            self.RRS.PowerRX=SNMP.GetPowerRXUL(self.Ip,self.Snmp_Community)
            self.RRS.Channel_Spasing=SNMP.GetCannel_SpacingUL(self.Ip,self.Snmp_Community)
            self.RRS.OppositeIp= SNMP.GetOppositeIPUL(self.Ip, self.Snmp_Community)
            self.RRS.Upper = SNMP.GetUpperUL(self.Ip, self.Snmp_Community)
            self.RRS.Band=RRL.GetBandFromFreq(self)
            self.RRS.Name=SNMP.GetNameUL(self.Ip,self.Snmp_Community)
            self.RRS.NameAsPL = NIOSS.RRNtoPL(self.RRS.Name)
            self.RRS.Polarization = "V"
            self.RRS.Reserv= "1+0"
            self.RRS.AmrMax,self.RRS.Capacity_max= SNMP.GetModulationUL(self.Ip,self.Snmp_Community)
        elif self.TypeIDU in ["OMNIBAS-2WCX"]:
            self.RRS.Name = SNMP.GetNameUL(self.Ip, self.Snmp_Community)
            n = self.getModemIndexOB(RRL_NIOSS)
            self.RRS.FreqTX = NIOSS.StripFreqOB(int(SNMP.GetFreqTXOBCX(self.Ip, self.Snmp_Community + '.' + n)) / 1000)
            self.RRS.FreqRX = NIOSS.StripFreqOB(int(SNMP.GetFreqRXOBCX(self.Ip, self.Snmp_Community + '.' + n)) / 1000)
            self.RRS.PowerTX = SNMP.GetPowerTXUL(self.Ip, self.Snmp_Community + '.' + n)
            self.RRS.PowerRX = SNMP.GetPowerRXUL(self.Ip, self.Snmp_Community + '.' + n)
            self.RRS.Channel_Spasing = SNMP.GetCannel_SpacingUL(self.Ip, self.Snmp_Community + '.' + n)
            self.RRS.OppositeIp = SNMP.GetOppositeIPUL(self.Ip, self.Snmp_Community + '.' + n)
            self.RRS.Upper = SNMP.GetUpperUL(self.Ip, self.Snmp_Community + '.' + n)
            self.RRS.Band = RRL.GetBandFromFreq(self)
            self.RRS.NameAsPL = NIOSS.RRNtoPL(self.RRS.Name)
            self.RRS.Polarization, self.RRS.Reserv = SNMP.GetPolariztionUL(self.Ip, self.Snmp_Community + '.' + n)
            self.RRS.AmrMax, self.RRS.Capacity_max = SNMP.GetModulationUL(self.Ip, self.Snmp_Community + '.' + n)
        elif self.TypeIDU in ["OmniBAS-8W","OmniBAS-4W","OmniBAS-2W",'OmniBAS-2W 16E1',"OmniBAS-4W v2","OmniBAS-4Wv2"]:
            self.RRS.Name = SNMP.GetNameUL(self.Ip, self.Snmp_Community)
            n=self.getModemIndexOB(RRL_NIOSS)
            self.RRS.FreqTX = NIOSS.StripFreqOB(int(SNMP.GetFreqTXUL(self.Ip, self.Snmp_Community+'.'+n))/1000)
            self.RRS.FreqRX = NIOSS.StripFreqOB(int(SNMP.GetFreqRXUL(self.Ip, self.Snmp_Community+'.'+n))/1000)
            self.RRS.PowerTX = SNMP.GetPowerTXUL(self.Ip, self.Snmp_Community+'.'+n)
            self.RRS.PowerRX = SNMP.GetPowerRXUL(self.Ip, self.Snmp_Community+'.'+n)
            self.RRS.Channel_Spasing = SNMP.GetCannel_SpacingUL(self.Ip, self.Snmp_Community+'.'+n)
            self.RRS.OppositeIp = SNMP.GetOppositeIPUL(self.Ip, self.Snmp_Community+'.'+n)
            self.RRS.Upper = SNMP.GetUpperUL(self.Ip, self.Snmp_Community+'.'+n)
            self.RRS.Band = RRL.GetBandFromFreq(self)
            self.RRS.NameAsPL = NIOSS.RRNtoPL(self.RRS.Name)
            self.RRS.Polarization,self.RRS.Reserv= SNMP.GetPolariztionUL(self.Ip, self.Snmp_Community+'.'+n)
            self.RRS.AmrMax, self.RRS.Capacity_max = SNMP.GetModulationUL(self.Ip, self.Snmp_Community+'.'+n)
        elif self.TypeIDU in ["iPasolink 200", "iPasolink 400", "iPasolink 1000", "iPasolink EX","iPasolink NEO"]:
            self.RRS.Name = SNMP.GetNameiPas(self.Ip, self.Snmp_Community)
            n = str(self.getModemIndexiPas(RRL_NIOSS))
            self.RRS.FreqTX = NIOSS.StripFreqOB(SNMP.GetFreqTXiPas(self.Ip, self.Snmp_Community,n))
            self.RRS.FreqRX = NIOSS.StripFreqOB(SNMP.GetFreqRXiPas(self.Ip, self.Snmp_Community,n))
            self.RRS.PowerTX = SNMP.GetPowerTXiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.PowerRX = SNMP.GetPowerRXiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.Channel_Spasing = SNMP.GetCannel_SpacingiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.OppositeIp = SNMP.GetOppositeIPiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.Upper = SNMP.GetUpperiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.Band = RRL.GetBandFromFreq(self)
            self.RRS.NameAsPL = NIOSS.RRNtoPL(self.RRS.Name)
            self.RRS.SubBand = SNMP.GetSubbandPas(self.Ip, self.Snmp_Community,n)
            self.RRS.Polarization, self.RRS.Reserv = SNMP.GetPolarizationiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.Capacity_max = SNMP.GetCapacityiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.AmrMin,self.RRS.AmrMax  = SNMP.GetModulationiPas(self.Ip, self.Snmp_Community,n)
        elif self.TypeIDU in ["iPasolink EX adv"]:
            self.RRS.Name = SNMP.GetNameiPas(self.Ip, self.Snmp_Community)
            n = str(self.getModemIndexiPasADV(RRL_NIOSS))
            self.RRS.FreqTX = NIOSS.StripFreqOB(SNMP.GetFreqTXiPas(self.Ip, self.Snmp_Community,n))
            self.RRS.FreqRX = NIOSS.StripFreqOB(SNMP.GetFreqRXiPas(self.Ip, self.Snmp_Community,n))
            self.RRS.PowerTX = SNMP.GetPowerTXiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.PowerRX = SNMP.GetPowerRXiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.Channel_Spasing = SNMP.GetCannel_SpacingiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.OppositeIp = SNMP.GetOppositeIPiPasADV(self.Ip, self.Snmp_Community,n)
            self.RRS.Upper = SNMP.GetUpperiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.Band = RRL.GetBandFromFreq(self)
            self.RRS.NameAsPL = NIOSS.RRNtoPL(self.RRS.Name)
            self.RRS.SubBand = SNMP.GetSubbandPas(self.Ip, self.Snmp_Community,n)
            self.RRS.Polarization, self.RRS.Reserv = SNMP.GetPolarizationiPas(self.Ip, self.Snmp_Community,n)
            self.RRS.Capacity_max = SNMP.GetCapacityiPasADV(self.Ip, self.Snmp_Community,n)
            self.RRS.AmrMin,self.RRS.AmrMax  = SNMP.GetModulationiPas(self.Ip, self.Snmp_Community,n)
        elif self.TypeIDU in ["iPasolink NEO ST"]:
            self.RRS.Name = SNMP.GetNameiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.FreqTX = NIOSS.StripFreqOB(SNMP.GetFreqTXiPasNEO(self.Ip, self.Snmp_Community))
            self.RRS.FreqRX = NIOSS.StripFreqOB(SNMP.GetFreqRXiPasNEO(self.Ip, self.Snmp_Community))
            self.RRS.PowerTX = SNMP.GetPowerTXiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.PowerRX = SNMP.GetPowerRXiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.Channel_Spasing = SNMP.GetCannel_SpacingiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.OppositeIp = SNMP.GetOppositeIPiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.Upper = SNMP.GetUpperiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.Band = RRL.GetBandFromFreq(self)
            self.RRS.NameAsPL = NIOSS.RRNtoPL(self.RRS.Name)
            self.RRS.SubBand = SNMP.GetSubbandPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.Polarization, self.RRS.Reserv = SNMP.GetPolarizationiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.Capacity_max = SNMP.GetCapacityiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.AmrMin,self.RRS.AmrMax  = SNMP.GetModulationiPasNEO(self.Ip, self.Snmp_Community)
        elif self.TypeIDU in ["iPasolink NEO HP"]:
            self.RRS.Name = SNMP.GetNameiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.FreqTX = NIOSS.StripFreqOB(SNMP.GetFreqTXiPasNEOHP(self.Ip, self.Snmp_Community))
            self.RRS.FreqRX = NIOSS.StripFreqOB(SNMP.GetFreqRXiPasNEOHP(self.Ip, self.Snmp_Community))
            self.RRS.PowerTX = SNMP.GetPowerTXiPasNEOHP(self.Ip, self.Snmp_Community)
            self.RRS.PowerRX = SNMP.GetPowerRXiPasNEOHP(self.Ip, self.Snmp_Community)
            self.RRS.Channel_Spasing = SNMP.GetCannel_SpacingiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.OppositeIp = SNMP.GetOppositeIPiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.Upper = SNMP.GetUpperiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.Band = RRL.GetBandFromFreq(self)
            self.RRS.NameAsPL = NIOSS.RRNtoPL(self.RRS.Name)
            self.RRS.SubBand = SNMP.GetSubbandPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.Polarization, self.RRS.Reserv = SNMP.GetPolarizationiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.Capacity_max = SNMP.GetCapacityiPasNEO(self.Ip, self.Snmp_Community)
            self.RRS.AmrMin,self.RRS.AmrMax  = SNMP.GetModulationiPasNEO(self.Ip, self.Snmp_Community)
        elif self.TypeIDU in ["AMM 2p B","AMM 6p С","AMM 6p C",'MINI-LINK 6692']:
            self.RRS.Name = SNMP.GetNameUL(self.Ip, self.Snmp_Community)
            n = self.getModemIndexML(RRL_NIOSS)
            self.RRS.FreqTX = NIOSS.StripFreqOB(SNMP.GetFreqTXML(self.Ip, self.Snmp_Community,n[4]))
            self.RRS.FreqRX = NIOSS.StripFreqOB(SNMP.GetFreqRXML(self.Ip, self.Snmp_Community,n[4]))
            self.RRS.PowerTX = SNMP.GetPowerTXML(self.Ip, self.Snmp_Community,n[4])
            self.RRS.PowerRX = SNMP.GetPowerRXML(self.Ip, self.Snmp_Community,n[4])
            self.RRS.Channel_Spasing = SNMP.GetCannel_SpacingML(self.Ip, self.Snmp_Community,n[4])
            self.RRS.OppositeIp = n[2]
            self.RRS.Upper = SNMP.GetUpper(self)
            self.RRS.Band = RRL.GetBandFromFreq(self)
            self.RRS.NameAsPL = NIOSS.RRNtoPL(self.RRS.Name)
            self.RRS.SubBand = SNMP.GetSubbandML(self.Ip, self.Snmp_Community,n[4])
            self.RRS.Polarization, self.RRS.Reserv = SNMP.GetPolarizationML(self.Ip, self.Snmp_Community,n[4])
            self.RRS.Capacity_max = SNMP.GetCapacityML(self.Ip, self.Snmp_Community,n[4])
            self.RRS.AmrMin,self.RRS.AmrMax  = SNMP.GetModulationML(self.Ip, self.Snmp_Community,n[4])
class iPas():
    def __init__(self):
        self.NumberIntToShort = {"1":"41","2":"42","3":"43","4":"44","5":"45","6":"46","7":"47","8":"48","9":"49","10":"50","11":"51","12":"52"}
        self.NumberIntToLong = {"1":"16842752","2":"25231360","3":"33619968","4":"42008576","5":"50397184","6":"58785792","7":"67174400","8":"75563008","9":"100728832","10":"109117440","11":"117506048","12":"125894656"}
        self.NumberIntToModem = {"1": "1", "2": "2", "3": "3", "4": "4", "5": "5", "6": "6", "7": "7", "8": "8",
                                 "9": "11", "10": "12", "11": "13", "12": "14"}

if __name__ == "__main__":
    myRRL = RRL()
    myRRL.ip = "10.42.163.50"
    myRRL.snmp_community = SNMP.get_community(myRRL.Ip)


