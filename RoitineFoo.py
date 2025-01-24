import requests
import re
import os
from Logging import LogThread, ToLog

def except_foo_dec(foo):
    def wrapper(*args, **kwargs):
        #ToLog(f"{foo.__name__}, *args = {args}, **kwargs = {kwargs} started")
        try:
            return foo(*args, **kwargs)
        except Exception as Err:
            ToLog(f"Error in {foo.__name__}, *args = {args}, **kwargs = {kwargs}, Error code = {Err}")
            raise Exception
        else:
            ToLog(f"{foo.__name__}, *args = {args}, **kwargs = {kwargs} finished successfully")
    return wrapper

# List of meetings
@except_foo_dec
def DataForMChanges(date = "03.12.2024", region = 0):
  
    # Подставляем дату в ссылку:
    ssilka = str("http://10.132.71.156/pls/ss/selector.sels.list?us=" + str(region) +
        "&str=" + date + "&wday=5")

    response = requests.get(
        f"http://10.132.71.156/pls/ss/selector.sels.list?us={region}&str={date}&wday=5")

    idsov = []
    namesov = []
    rezhimsov = []
    timesov = []
    initsov = []
    studiasov = []
    themesov = []
    spisoksov = []
    spisokuchast = []
    nomsov = 1

    # Если Управление
    if int(region) == 0:
        
        filesplit = response.text.splitlines()
        #print(str(filesplit))

        #Задаем параметры поиска
        poisk = "&us=0&sid="
        poisk2 = '''<td width=15% class=zag>Примечание</td>'''
        poisk3 = '''&nbsp;</td></tr>'''
        poisk4 = '''<td class="zag" rowspan=2>'''
        poisk5 = '''<td class="msk" rowspan=2>'''
        poisk6 = '''<a href="javascript:go(0,1,0'''
        poisk7 = '''Регион-'''
        poisk8 = '''&nbsp;'''
        poisk9 = '''&nbsp;</td><td class=norm>&nbsp'''
        nachalo = "2"
        konec = '''</td>'''

        #Обработка кода страницы и составление списков
        for i in range(0, len(filesplit)-1):
            filesplit[i] = str(filesplit[i]).strip()

            #добавление в списки разделителей - строки Начало совещания и Список участников
            if (
                filesplit[i].find('''<td width=15% class=zag>Примечание</td>''')>-1
                or
                filesplit[i]=='''&nbsp;</td></tr>'''):
                spisoksov.append(str(nomsov))
                spisokuchast.append("Список участников  "+str(nomsov))
                nomsov = nomsov + 1

            #составление списка участников конференций (необработанного)
            if filesplit[i].find('''<a href="javascript:go(0,1,0''')!=-1:
                spisokuchast.append(filesplit[i][filesplit[i].find('''">''')+2:filesplit[i].find('''</a>''')])
 
            #составление списка как в SMS
            if (
                (filesplit[i].find(poisk4)!=-1)
                or
                (filesplit[i].find(poisk5)!=-1)):
                if filesplit[i].find('''<br>''') > -1:
                    filesplit[i] = filesplit[i][:filesplit[i].find('''</td>''')+1]
                filesplit[i] = filesplit[i][filesplit[i].find(nachalo) + 2:filesplit[i].find(konec)]
                spisoksov.append(filesplit[i])
                #print("\tfrom SMS = " + str(filesplit[i]))

            #find themes
            if filesplit[i].find(poisk9)!=-1:
                filesplit[i] = filesplit[i][filesplit[i].find(poisk9) + len(poisk9) + 1:]
                spisoksov.append(filesplit[i])
                #print("\ttheme = " + str(filesplit[i]))

            #составление списка ID конференций (внутри списка SMS)
            if filesplit[i].find(poisk)!=-1:
                filesplit[i] = filesplit[i][filesplit[i].find(poisk)+10:filesplit[i].find('''>"''')-1]
                spisoksov.append(str(filesplit[i]))

        #print("begin of deparse")
        for i in range (6, len(spisoksov)):
            #print(str(spisoksov[i]))
            if (i+1)%7==0:
                studiasov.append(spisoksov[i-5])
                rezhimsov.append(spisoksov[i-4])
                timesov.append(spisoksov[i-3])
                initsov.append(spisoksov[i-2])
                themesov.append(spisoksov[i-1])
                idsov.append(spisoksov[i])

    # формируем списки с учетом отмен и проверок
    idsov1 = []
    studiasov1 = []
    themesov1 = []
    #namesov1 = []
    rezhimsov1 = []
    initsov1 = []
    timesov1 = []
    uchastsov1 = []
    #nomer = []
    #nomernach = 1
    uchastsov = []
    temp_uchast = []

    # Формируем список списков участников
    for i in range (1, len(spisokuchast)):
        if spisokuchast[i].find("Список участников")==-1:
            temp_uchast.append(spisokuchast[i])
        else:
            if len(temp_uchast)==0:
                uchastsov.append(["None"])
                temp_uchast.clear()
            else:
                uchastsov.append(temp_uchast[:])
                temp_uchast.clear()
                    
    for i in range (0, len(timesov)):
        if 1 == 0:
            pass
        else:        
            if len(timesov[i]) < 16:
                timesov[i] = timesov[i].replace("<br>-<br>", "")
                if timesov[i][-1] == ":":
                    timesov[i] = timesov[i][:-1]
                timesov1.append(timesov[i])
                             
            else:
                timesov1.append(timesov[i].replace("<br>-<br>", "-"))

    itog = {}
    for item in range (0, len(idsov)):
        itog.update(
            {idsov[item]: [idsov[item], studiasov[item], initsov[item],
                           timesov1[item], rezhimsov[item], themesov[item],
                           ", ".join(uchastsov[item][:])]})
 
    return itog

#===============================================
#===============================================
#===============================================
#=============================================== 
# Making list of meeting with some id
@except_foo_dec
def DataOneMeeting(idsov):
    itogitogov = []
    for k in range (0,5):
        ssilkaSS = str(
            "http://10.132.71.156/pls/ss/selector.report.study_p?sid=" + idsov + "&us=" + str(k))
        
        responseSS = requests.get(ssilkaSS)
        filesplit = responseSS.text.splitlines()

        #Задаем наши списки и номер совещания
        dolgnost = []
        fio = []
        prim = []

        #Обработка кода страницы и составление списков
        for i in range(0, len(filesplit)-1):                            
                # Формируем списки для таблицы - должность, ФИО, Примечание
            if (
                filesplit[i].find('''<tr><td colspan=3 class=z2>''')!=-1
                or
                filesplit[i].find('''<tr><td class=spr valign=top>''')!=-1):
                
                if filesplit[i].find('''<tr><td colspan=3 class=z2>''')!=-1:
                    dolgnost.append("КАБИНЕТ" + str(filesplit[i][filesplit[i].find('''z2>''')+3:filesplit[i].find('''</td>''')]))
                    fio.append("NONE")
                    prim.append("NONE")

                if filesplit[i].find('''<tr><td class=spr valign=top>''')!=-1:
                    dolgnost.append(filesplit[i+1])
                    fio.append("Новые участники:")
                    n = i
                    while filesplit[n].find('''</table></td>''')==-1:
                        n = n+1
                    for s in range (i,n):
                       
                        if filesplit[s].find('''<td class=spr>''')!=-1:
                            fio.append(
                                str(filesplit[s])[filesplit[s].find('''<td class=spr>''')+14:filesplit[s].find('''&nbsp;&nbsp''')]+
                                str("  ")+str(filesplit[s])[filesplit[s].find('''&nbsp;&nbsp''')+12:filesplit[s].find('''</td>''')])
                        elif (
                            filesplit[s+1].find('''</table></td>''')!=-1
                            and
                            filesplit[s].find('''<table width=''')!=-1):
                            fio.append("PUSTO")
                                                
                    prim.append(filesplit[n+2])

        # Преобразовываем список участников, чтобы сгруппировать их по должностям    
        fio.append(str("Новые участники"))
        fio1 = []

        for i in range (0, len(fio)):
            if fio[i] =="NONE":
                fio1.append("NONE")
            elif (fio[i].find("Новые участники")!=-1) and (i<(len(fio)-2)):
                fio1.append(" ")
                n = i+1
                while fio[n].find("Новые участники")==-1:
                    n = n+1
                for s in range (i,n):
                    if fio[s].find("Новые участники")==-1:
                        temp = fio1[-1][:]
                        fio1[-1] = temp+"/NEXT/"+fio[s][:]

        #for i in range(0, len(prim)):
        #    print ("Долж = "+dolgnost[i]+" --- ФИО = "+fio1[i]+" --- Прим = "+prim[i])

        # итог цикла
        itog = [dolgnost, fio1, prim]
        itogitogov.append(itog)

    itoglist = {}
    for k in range (0, 5):
        dolg = itogitogov[k][0]
        fio = itogitogov[k][1]
        prim = itogitogov[k][2]

        for i in range (1, len(dolg)+1):
            dolg[i-1] = dolg[i-1].replace("&nbsp;", " ").strip()
            fio[i-1] = fio[i-1][7:].replace("/NEXT/"," ")
            fio[i-1] = fio[i-1].replace("PUSTO", " ")
            fio[i-1] = fio[i-1].replace("NONE", "").strip()
            prim[i-1] = prim[i-1].replace("&nbsp;"," ").strip()

    
        itoglist.update({str(k): [dolg[:], fio[:], prim[:]]})

    result = []
    cab = "ERROR CAB"
    for key in itoglist.keys():
        if len(itoglist[key][0]) > 0:
            for dolg in range (0, len(itoglist[key][0])):
               txt = itoglist[key][0][dolg]
               if txt.find("КАБИНЕТ") != -1:
                   cab = txt[txt.find("КАБИНЕТ") + 7 :]
                   #print(f"new cab = {cab}")
               else:
                   result.append((key, cab, itoglist[key][0][dolg], itoglist[key][1][dolg], itoglist[key][2][dolg]))      
               
    return result
    
