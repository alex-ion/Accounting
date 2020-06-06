import time, xlsxwriter, os, shutil, smtplib, zipfile, psycopg2
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import timedelta, date, datetime

DB_connected = False
eroare_alerta_mail=''

try:
    db = psycopg2.connect(database='___',
      user='___',
      password='___',
      host='___',
      port='___'
    )
    cur = db.cursor()
    DB_connected = True
except(Exception) as error:
    eroare_alerta_mail += str(error)+'<BR>'

eroare_alerta_mail=''
RawText=''
listacod=["Contract","Card","Expiration","Suma","Autorizare","Client","CI_Date","Brand"]#lista cu tipurile de date de pe fiecare linie din fisierul de interpretat
Lista_Barclays=[]
Lista_Refund=[]
Lista_AMEX=[]
Lista_Diners_Discover=[]
Row=1#reprezinta randul de pe care incepem sa scriem datele
Column=0#reprezinta coloana de pe care incepem sa scriem datele
locatie_fisiere_Avis=""
locatie_fisiere_Budget=""
locatie_arhive_Avis=""
locatie_arhive_Budget=""

class Contract():
    def __init__(self,contract,card,expiration,suma,autorizare,client,CI_Date,Brand):
        self.Contract=contract.strip("E")
        self.Card=card
        self.Expiration=expiration[0]+expiration[1]+"/"+expiration[2]+expiration[3]
        self.Autorizare=autorizare
        self.Client=client.strip()
        self.Brand=Brand
        self.Refund='false'
        
        #Aici transformam suma dintr-o insiruire de cifre intr-un numar cu 2 zecimale
        suma_temp=''
        j=0
        for i in suma:
            if j<len(suma)-5:
                suma_temp=suma_temp+suma[j]
                j+=1
        if suma[len(suma)-5]==" ":
            self.Suma=suma_temp+'.'+"0"+suma[len(suma)-4]+suma[len(suma)-3]+suma[len(suma)-2]+suma[len(suma)-1]
        else:
            self.Suma=suma_temp+'.'+suma[len(suma)-5]+suma[len(suma)-4]+suma[len(suma)-3]+suma[len(suma)-2]+suma[len(suma)-1]
            
        CI_Date=time.strptime("20"+CI_Date,'%Y/%m/%d')
        self.CI_Date=str(CI_Date.tm_year)+"-"+str(CI_Date.tm_mon).rjust(2,"0")+"-"+str(CI_Date.tm_mday).rjust(2,"0")
        if self.Card[0]=='6':
            self.Card=self.Card[0]+self.Card[1]+self.Card[2]+self.Card[3]+self.Card[4]+self.Card[5]+'XXXXXX'+self.Card[12]+self.Card[13]+self.Card[14]+self.Card[15]
        if "CR" not in self.Suma:
            if self.Card.startswith("37") or self.Card.startswith("34"):
                Lista_AMEX.append(self)
                self.Procesator = 'AMEX'
                self.Identificator = 'CA'
            elif self.Card.startswith("36"):
                Lista_Diners_Discover.append(self)
                self.Procesator = 'DINERS'
                self.Identificator = 'CD'
            elif self.Card.startswith("6"):
                Lista_Diners_Discover.append(self)
                self.Procesator = 'DINERS'
                self.Identificator = 'CS'
            elif self.Card.startswith("5"):
                Lista_Barclays.append(self)
                self.Procesator = 'BARCLAYS'
                self.Identificator = 'CM'               
            elif self.Card.startswith("4"):
                Lista_Barclays.append(self)
                self.Procesator = 'BARCLAYS'
                self.Identificator = 'CX'               
        else:
            self.Suma=self.Suma.strip("CR").strip(" ")
            Lista_Refund.append(self)
            if self.Card.startswith("37") or self.Card.startswith("34"):
                self.Procesator = 'AMEX'
                self.Identificator = 'CA'
                self.Refund='true'
            elif self.Card.startswith("36"):
                self.Procesator = 'DINERS'
                self.Identificator = 'CD'
                self.Refund='true'                
            elif self.Card.startswith("6"):
                self.Procesator = 'DINERS'
                self.Identificator = 'CS'
                self.Refund='true'                
            elif self.Card.startswith("5"):
                self.Procesator = 'BARCLAYS'
                self.Identificator = 'CM'
                self.Refund='true'                
            elif self.Card.startswith("4"):
                self.Procesator = 'BARCLAYS'
                self.Identificator = 'CX'  
                self.Refund='true'

                
def scriere_log(mesaj): #scrie mesajul in log si da mail daca al doilea parametru este setat pe True
    LogFile=open("LogFileConta4.txt","a")
    LogFile.write(str(time.ctime())+": "+mesaj)
    LogFile.close()
    print (str(time.ctime())+mesaj)

def impartire_in_liste(brand):
    global Row, Lista_Barclays, Lista_Refund, Lista_AMEX, Lista_Diners_Discover, RawText
    Row=1
    for e in RawText:
        if ("CREDIT CLUB SYSTEM" in e)\
           or ("ROMANIA RENTALS PROCESSED CENTRALLY " in e)\
           or ("CCAA06E /01 RO" in e)\
           or ("CREDIT CLUB  CA   AMEX" in e)\
           or ("CREDIT CLUB  CM   MCARD" in e)\
           or ("CREDIT CLUB  CX   VISA" in e)\
           or ("                                                                                                                                     \n" in e)\
           or ("                                                                                                                     CURR           " in e)\
           or ("CHECKIN LOCATION     RA NUMBER      ACCT NUMBER     EXP    RENT AMT    AUTH           CUST NAME         BILL AMT   CODE CI DATE" in e)\
           or ("         *** NO DETAILS FOR THIS REPORT ****          " in e)\
           or ("******************************************************" in e)\
           or ("*                    END OF REPORT                   *" in e)\
           or ("" in e)\
           or ("                                                              " in e):
            pass
        else:
            j=24 #reprezinta al 24-lea caracter de pe linia de unde incepem sa copiem datele din fisier
            rand=""
            while j<len(e):
                rand=rand+e[j]
                j+=1
            obiect = len(globals())
            globals()[obiect] = Contract(rand[0:10],rand[12:28],rand[31:35],rand[38:48],rand[49:55],rand[60:80],rand[99:107],brand)                 

def generare_raport(parametru1,parametru2,parametru3):
    global Row, Column, workbook, worksheet, merge_format, formattitlu, formatcontinut, formatcontinut2
    parametru4=0
    sir="A"+str(Row)+":H"+str(Row)
    worksheet.merge_range(sir,parametru1, merge_format)
    for e in listacod:
        worksheet.write(Row,Column,e,formattitlu)#creare cap de tabel in Excel
        Column+=1
    Column=0
    Row+=1
    for e in parametru2:
        worksheet.write(Row,Column,int(e.Contract),formatcontinut)
        worksheet.write(Row,Column+1,e.Card,formatcontinut)
        worksheet.write(Row,Column+2,e.Expiration,formatcontinut)
        worksheet.write(Row,Column+3,float(e.Suma),formatcontinut)
        worksheet.write(Row,Column+4,e.Autorizare,formatcontinut)
        worksheet.write(Row,Column+5,e.Client,formatcontinut2)
        worksheet.write(Row,Column+6,e.CI_Date,formatcontinut)
        worksheet.write(Row,Column+7,e.Brand,formatcontinut)
        Row+=1
        parametru4=parametru4+float(e.Suma)
    worksheet.write(Row,2,parametru3,formattitlu)#aici scriem numele totalului
    worksheet.write(Row,3,parametru4,formattitlu)#aici scriem totalul sumelor
    Row=Row+4

def alerta_mail(text):
    msg = MIMEMultipart()
    fromaddr = "___"
    toaddr = "___"
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = "Erori incasari Avis&Budget "
    body = str(text)
    part = MIMEBase('application', 'octet-stream')
    msg.attach(MIMEText(body, 'html'))
    server = smtplib.SMTP_SSL('___', 465)
    server.login("___", "___")
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()
    time.sleep(2)
    

def trimitere_email(nume_fisier_procesat,zi,refund,fara_incasari):
    fromaddr = "___"
    toaddr = "___"
    if refund:
        toccaddr = "___,___"
    else:
        toccaddr = "___"
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['CC'] = toccaddr
    msg['Subject'] = "Incasari Avis&Budget "+zi
    if fara_incasari == True:
        body = "In ziua de "+zi+" nu au fost incasari.<BR><BR><BR> Acest mail a fost generat automat, va rugam nu dati reply."
    else:
        body = "Fisierul cu incasarile Avis&Budget din "+zi+" este disponibil in atasament.<BR><BR><BR> Acest mail a fost generat automat, va rugam nu dati reply."
        attachment = open(os.path.join(str(os.getcwd()),nume_fisier_procesat), "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % nume_fisier_procesat)
        msg.attach(part)
    msg.attach(MIMEText(body, 'html'))
    server = smtplib.SMTP_SSL('___, 465)
    server.login("___", "___")
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr.split(",") + toccaddr.split(","), text)
    server.quit()
    time.sleep(2)

def main(fisier_Avis,fisier_Budget,zi):
    global RawText, workbook, worksheet, merge_format, formattitlu, formatcontinut, formatcontinut2, Lista_Barclays, Lista_Refund, Lista_AMEX, Lista_Diners_Discover
    Lista_Barclays=[]
    Lista_Refund=[]
    Lista_AMEX=[]
    Lista_Diners_Discover=[]
    fara_incasari = False
    try:
        RawText=open(fisier_Avis,"r") #Incearca deschiderea fisierului Avis care trebuie procesat
        impartire_in_liste("Avis")
        RawText.close()
        RawText=open(fisier_Budget,"r") #Incearca deschiderea fisierului Budget care trebuie procesat
        impartire_in_liste("Budget")
        RawText.close()
        nume_fisier_procesat = "CCAA06EQ-RUN  ON   "+zi+".xlsx"
        workbook=xlsxwriter.Workbook(nume_fisier_procesat)
        worksheet=workbook.add_worksheet()
        worksheet.set_column('A:B', 18)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:E', 18)
        worksheet.set_column('F:F', 25)
        worksheet.set_column('G:H', 18)
        formattitlu = workbook.add_format({'align': 'center','bold': True})
        formatcontinut = workbook.add_format({'align': 'center'})
        formatcontinut2 = workbook.add_format({'align': 'left'})
        merge_format = workbook.add_format({'bold': 1,'border': 1,'align': 'center','valign': 'vcenter'})
        if Lista_Barclays:
            generare_raport("BARCLAYS",Lista_Barclays, "Total Barclays")
        if Lista_Diners_Discover:
            generare_raport("DINERS & DISCOVER",Lista_Diners_Discover, "Total Diners & Discover")                
        if Lista_AMEX:
            generare_raport("AMEX",Lista_AMEX,"Total AMEX")
        if Lista_Refund:
            generare_raport("REFUND",Lista_Refund,"Total Refund")
        workbook.close()
        if (Lista_Barclays == []) &(Lista_Refund == [])& (Lista_AMEX == []) & (Lista_Diners_Discover == []):
            fara_incasari = True
        trimitere_email(nume_fisier_procesat,zi,Lista_Refund,fara_incasari)
        shutil.copy(nume_fisier_procesat,os.path.join(locatie_fisiere_Avis,nume_fisier_procesat))
        shutil.move(nume_fisier_procesat,os.path.join(locatie_fisiere_Budget,nume_fisier_procesat))
        scriere_log(": S-a procesat cu succes fisierul Avis&Budget '"+nume_fisier_procesat+"'\n")
        for lista in [Lista_Barclays,Lista_Diners_Discover,Lista_AMEX,Lista_Refund]:
            for element in lista:
                inserare_in_DB(element,"CCAA06EQ-RUN  ON   "+zi+".txt")
    except(IOError) as E:
        scriere_log("A aparut eroarea: "+str(E).upper())

def inserare_in_DB(obiect,nume_fisier):
    global eroare_alerta_mail
    query = "INSERT INTO alex.incasari VALUES (default,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}',default)".format(obiect.Contract,obiect.Card,obiect.Expiration,obiect.Identificator,obiect.Procesator,obiect.Suma.strip(),obiect.Refund,obiect.Autorizare,obiect.Client,obiect.CI_Date,obiect.Brand,nume_fisier)
    try:
        cur.execute(query)
        db.commit()
    except(Exception) as error:
        eroare_alerta_mail += str(error)+'<BR>'

def daterange(date1, date2):
    for n in range(int ((date2 - date1).days)+1):
        yield date1 + timedelta(n)


        
if DB_connected == True:
    #Daca cumva scriptul nu a rulat intr-o zi/cateva zile la rand, cu prima ocazie cand ruleaza verifica in baza de date ce zile nu au rulat si insereaza default False pe ele
    #preluare lista de zile din baza de date
    query = "SELECT zi_calendaristica FROM alex.zile_rulate"
    lista_zile_DB=[]
    try:
        cur.execute(query)
        db.commit()
        for row in cur.fetchall():
            lista_zile_DB.append(row[0])
    except(Exception) as error:
        eroare_alerta_mail += str(error)+'<BR>'

    for zi in daterange(date(2020, int('03'), int('01')), date(time.localtime().tm_year, time.localtime().tm_mon, time.localtime().tm_mday)):
        if str(zi) not in lista_zile_DB:
            zi_din_an=str(time.strptime(str(zi),"%Y-%m-%d").tm_yday).rjust(3,"0")
            query = "INSERT INTO alex.zile_rulate VALUES (default, '{0}','{1}','false',default)".format(zi_din_an,str(zi))    
            try:
                cur.execute(query)
                db.commit()
            except(Exception) as error:
                eroare_alerta_mail += str(error)+'<BR>'

    #se preiau toate inregistrarile din tabele si se verifica pentru ce zile avem False
    zile_de_procesat=[]
    query = "SELECT * FROM alex.zile_rulate"
    try:
        cur.execute(query)
        db.commit()
        for row in cur.fetchall():
            if row[3]=='false':
                zile_de_procesat.append(row[2])
    except(Exception) as error:
        eroare_alerta_mail += str(error)+'<BR>'

    #interogare fisier zile-rulate, pentru intocmirea listei de zile in care nu s-a rulat raportul
    for zi in zile_de_procesat:
    ##    zi_formatata = str(time.strptime("2020-04-12","%Y-%m-%d").tm_mday) + str(time.strptime("2020-04-12","%Y-%m-%d").tm_mon) + str(time.strptime("2020-04-12","%Y-%m-%d").tm_year)
        zi_formatata=datetime.strftime(datetime.strptime(zi,"%Y-%m-%d"),"%d%b%y")
        zi_din_an=str(time.strptime(zi,"%Y-%m-%d").tm_yday).rjust(3,"0")
        try:
            zip=zipfile.ZipFile(os.path.join(locatie_arhive_Avis,"eproccda{0}.zip".format(zi_din_an)))
            zip.extractall(os.path.join(locatie_fisiere_Avis,"TEMP"))
            shutil.move(os.path.join(os.path.join(locatie_fisiere_Avis,"TEMP"),"CCAA06EQ"),os.path.join(locatie_fisiere_Avis,"CCAA06EQ-RUN  ON   {0}.txt".format(zi_formatata)))
        except(Exception) as error:
            eroare_alerta_mail += str(error)+'<BR>'
        try:
            zip=zipfile.ZipFile(os.path.join(locatie_arhive_Budget,"bud_eproccda{0}.zip".format(zi_din_an)))
            zip.extractall(os.path.join(locatie_fisiere_Budget,"TEMP"))
            shutil.move(os.path.join(os.path.join(locatie_fisiere_Budget,"TEMP"),"CCAA06EQ"),os.path.join(locatie_fisiere_Budget,"CCAA06EQ-RUN  ON   {0}.txt".format(zi_formatata)))
        except(Exception) as error:
            eroare_alerta_mail += str(error)+'<BR>'
            
        try:
            zip=zipfile.ZipFile(os.path.join(locatie_arhive_Avis,"eproccdb{0}.zip".format(zi_din_an)))
            zip.extractall(os.path.join(locatie_fisiere_Avis,"TEMP"))
            shutil.move(os.path.join(os.path.join(locatie_fisiere_Avis,"TEMP"),"CCAA06ES"),os.path.join(locatie_fisiere_Avis,"CCAA06ES-RUN  ON   {0}.txt".format(zi_formatata)))
        except(Exception) as error:
            eroare_alerta_mail += str(error)+'<BR>'
                           
        try:
            zip=zipfile.ZipFile(os.path.join(locatie_arhive_Budget,"bud_eproccdb{0}.zip".format(zi_din_an)))
            zip.extractall(os.path.join(locatie_fisiere_Budget,"TEMP"))
            shutil.move(os.path.join(os.path.join(locatie_fisiere_Budget,"TEMP"),"CCAA06ES"),os.path.join(locatie_fisiere_Budget,"CCAA06ES-RUN  ON   {0}.txt".format(zi_formatata)))
        except(Exception) as error:
            eroare_alerta_mail += str(error)+'<BR>'

        nume_fisier="CCAA06EQ-RUN  ON   "+str(zi_formatata)+".txt"
        fisier_Avis=os.path.join(locatie_fisiere_Avis,nume_fisier)
        fisier_Budget=os.path.join(locatie_fisiere_Budget,nume_fisier)
        if os.path.isfile(fisier_Avis) and os.path.isfile(fisier_Budget):
            scriere_log(": Exista fisierul '"+nume_fisier+"' atat pentru Avis cat si pentru Budget\n")
            main(fisier_Avis,fisier_Budget,str(zi_formatata))
            query = "UPDATE alex.zile_rulate SET rulat='true' WHERE zi_calendaristica='{0}'".format(zi)
            try:
                cur.execute(query)
                db.commit()
            except(Exception) as error:
                eroare_alerta_mail += str(error)+'<BR>'
        elif os.path.isfile(fisier_Avis) and not os.path.isfile(fisier_Budget):
            scriere_log(": Nu exista fisierul Budget '"+nume_fisier+"'\n")
        elif not os.path.isfile(fisier_Avis) and os.path.isfile(fisier_Budget):
            scriere_log(": Nu exista fisierul Avis '"+nume_fisier+"'\n")
        elif not os.path.isfile(fisier_Avis) and not os.path.isfile(fisier_Budget):
            scriere_log(": Nu exista fisierul '"+nume_fisier+"' nici pentru Avis nici pentru Budget\n")
            

    db.close()
if eroare_alerta_mail:
    alerta_mail(eroare_alerta_mail)
