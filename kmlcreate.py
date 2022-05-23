import simplekml
import xlrd

if __name__ == "__main__":
    kml = simplekml.Kml()
    bs = []
    print("введите РРЛ для подгтовки файла с пролетами и МВН для подготовки файла с узлами")
    str_value=input()
    if str_value=="РРЛ":
        rb = xlrd.open_workbook('C:/Users/yvsobol3/Documents/Список с адресами.xlsm')
        sheet = rb.sheet_by_index(0)
        folder2020 = kml.newfolder(name='2020')
        folder2021 = kml.newfolder(name='2021')
        folder2022 = kml.newfolder(name='2022')
        folderother = kml.newfolder(name='other')
        for rownum in range(sheet.nrows):
            row = sheet.row_values(rownum)

            if row[0] and str(row[58]).lower()=="да":
                if (row[49]==0) or (row[49]==""):
                    folder = folderother
                    if not isinstance(row[43], str):
                        date = xlrd.xldate.xldate_as_datetime(row[43], rb.datemode).date()
                        if date.year == 2020:
                            folder=folder2021
                        elif date.year == 2021:
                            folder = folder2021
                        elif date.year == 2022:
                            folder = folder2022
                    else:
                        date = row[43]
                    rrl = folder.newlinestring(name=row[0]+''+row[1], coords=[(row[55], row[54]),(row[57], row[56])],
                                              description=str(row[4]) + '\n' + str(date))
                    if row[0] not in bs:
                        bs.append(row[0])
                        rrsA = folder.newpoint(name=row[0], coords=[(row[55], row[54])])
                        rrsA.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
                    if row[1] not in bs:
                        bs.append(row[1])
                        rrsB = folder.newpoint(name=row[1], coords=[(row[57], row[56])])
                        rrsB.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
                    c_Color = 'ffffffff'
                    if  'RTN' in str(row[4]):
                        c_Color = 'ffd4ff7f'  # желтый
                    elif ('GX' in str(row[4])) or  ('FX' in str(row[4])):
                        c_Color = 'ff9afa00'  # салатовый
                    elif ('mini' in str(row[4]).lower()) or  ('ml' in str(row[4]).lower()):
                        c_Color = 'ff808000'  # темнобирюзовый
                    elif ('omni' in str(row[4]).lower()) or  ('ob' in str(row[4]).lower()):
                        c_Color = 'ffcd5ba9'  # фиолетовый
                    elif 'ipas' in str(row[4]).lower():
                        c_Color = 'ff60a4f4'  # оранжевый


                    rrl.style.linestyle.color = c_Color
                    rrl.style.linestyle.width = 3
        #
        #
        #
        #
        #
        kml.save("MW.kml")

    elif str_value=="MBH ALL":
        rb = xlrd.open_workbook('C:/Users/yvsobol3/Downloads/prl_rack___7011.xlsx')
        sheet = rb.sheet_by_index(0)
        folderExist = kml.newfolder(name='СУЩ')
        folderPlanning = kml.newfolder(name='ПЛАН')
        for rownum in range(sheet.nrows):
            row = sheet.row_values(rownum)
            if (row[9]!=0) and (row[9]!=""):
                if row[8]==100:
                    folder=folderPlanning
                    c_Color = 'ffd4ff7f'  # желтый
                elif row[8]==0:
                    folder = folderExist
                    c_Color = 'ff9afa00'  # салатовый
                else:
                    continue

                pnt = folder.newpoint(name=str(row[9]), coords=[(row[13], row[12])],
                                      description=str(row[11]))
                pnt.style.iconstyle.color =c_Color
                pnt.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pal3/icon52.png'



        #
        #
        #
        #
        kml.save("MBH ALL.kml")
    elif str_value == "MBH":
        rb = xlrd.open_workbook('M:/_ОтделРазвитияТранспортныхСетей/!_Функциональная_группа_DWDM_MBH_SDH/2017_MBH(план_исполн_закр.док)/!ПЛАН!/План MBH общий.xlsm')
        sheet = rb.sheet_by_index(0)
        folder2020 = kml.newfolder(name='2020')
        folder2021 = kml.newfolder(name='2021')
        folder2022 = kml.newfolder(name='2022')
        folderother = kml.newfolder(name='other')
        for rownum in range(sheet.nrows):
            row = sheet.row_values(rownum)
            if row[0]:
                if str(row[10]).lower()=="да":
                    folder = folderother
                    if not isinstance(row[9], str):
                        date = xlrd.xldate.xldate_as_datetime(row[9], rb.datemode).date()
                        if date.year == 2020:
                            folder=folder2021
                        elif date.year == 2021:
                            folder = folder2021
                        elif date.year == 2022:
                            folder = folder2022
                    else:
                        date = row[9]

                    pnt = folder.newpoint(name=row[18], coords=[(row[38], row[37])],
                                              description=str(row[3]) + '\n' + str(date))
                    pnt.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pal3/icon52.png'
                    if  ('8603' in str(row[3])) or ('8615' in str(row[3]))or ('8609' in str(row[3]))or ('8611' in str(row[3])):
                        pnt.style.iconstyle.color = 'ff2fffad'  # greenyellow
                    elif ('8625' in str(row[3])) or ('8665' in str(row[3])):
                        pnt.style.iconstyle.color =  'ffff0000'  # blue
                    elif 'модернизация' in str(row[2]):
                        pnt.style.iconstyle.color = 'ff0000ff' #yellow



        kml.save("MBH.kml")