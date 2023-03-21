# -*- coding: utf-8 -*- 
import os
import sys
import time
import shutil

from reportlab import *
from reportlab.lib.utils import *
from reportlab.lib.colors import CMYKColor, PCMYKColor


############ Formato A4: 595 x 842 ############

############ Posiziono il file di log.. ############
Anno_Mese = time.strftime("%Y_%m")

SoggettoEmail = ""
try:
    f_log = open (".\\" + Anno_Mese + "_TARROS_QTZ.log",'a')
except IOError:
    sys.exit (1)
    
f_log.write ("--------------------------------------------------------------------------------\n")
f_log.write (time.asctime(time.localtime()))
f_log.write (" Inizio Elaborazione\n")

lista = sys.argv
for i in lista:
    stringa_log = time.asctime(time.localtime()) + " "+ i + "\n"
    f_log.write (stringa_log)
if len(sys.argv)!= 2:
    f_log.write (time.asctime(time.localtime()))
    f_log.write (" Numero di parametri errato..\n")
    sys.exit(1)

############ Apro e leggo il file INI per la postalizzazione ############
try:
    f_ini = open("Tarros.ini",'r')
    #print ("File ini",f_ini.name, "aperto..")
except IOError:
    f_log.write (time.asctime(time.localtime()))
    f_log.write (" Impossibile aprire il file ini: Tarros.ini\n")
    sys.exit (2)

for linea in f_ini:
    #print (linea,end='')
    base_linea = linea.split ("=")
    if base_linea[0] == "MAIL_TO":
        MAIL_TO = base_linea[1]
        MAIL_TO = MAIL_TO.rstrip('\n')
        if len (MAIL_TO) == 0:
            MAIL_TO = "stefano.zanni@tarros.it"
    if base_linea[0] == "MAIL_USER":
        MAIL_USER = base_linea[1]
        MAIL_USER = MAIL_USER.rstrip('\n')
    if base_linea[0] == "MAIL_PWD":
        MAIL_PWD = base_linea[1]
        MAIL_PWD = MAIL_PWD.rstrip('\n')
    if base_linea[0] == "MAIL_HOST":
        MAIL_HOST = base_linea[1]
        MAIL_HOST = MAIL_HOST.rstrip('\n')
    if base_linea[0] == "MAIL_PORT":
        MAIL_PORT = base_linea[1]
        MAIL_PORT = MAIL_PORT.rstrip('\n')
    if base_linea[0] == "DIR_RIFERIMENTO":
        DIR_RIFERIMENTO = base_linea[1]
        DIR_RIFERIMENTO = DIR_RIFERIMENTO.rstrip('\n')
    if base_linea[0] == "AUTENTICAZIONE":
        AUTENTICAZIONE = base_linea[1]
        AUTENTICAZIONE = AUTENTICAZIONE.rstrip('\n')
        #f_log.write (time.asctime(time.localtime()))
        #f_log.write (" AUTENTICAZIONE: " + str(AUTENTICAZIONE) + "\n")

                
f_ini.close() 
POSTA = "NO"

############ Cancello i vecchi dta processati nella dir di out della printer NEXTFILE dello spooler

processed_dir = DIR_RIFERIMENTO + "\\Out\\"
for f in os.listdir(processed_dir):
    if f.endswith(".processed"):
        os.remove(os.path.join( processed_dir + f))
        #f_log.write (time.asctime(time.localtime()))
        #f_log.write (" Elimino il file " + str(f) + " dalla dir " + DIR_RIFERIMENTO + "\\Out\n")


############ IMPORT E DEFINIZIONI PER INVIO EMAIL DI NOTIFICA ############

import smtplib
import mimetypes
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.MIMEAudio import MIMEAudio
from email.MIMEImage import MIMEImage
from email.Encoders import encode_base64

def sendMail(subject, text, *attachmentFilePaths):
    User = MAIL_USER
    Password = MAIL_PWD
    #recipient = MAIL_TO
    recipient = indirizzoEmail

    msg = MIMEMultipart()
    msg['From'] = User
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    for attachmentFilePath in attachmentFilePaths:
        msg.attach(getAttachment(attachmentFilePath))
    
    try:
        mailServer = smtplib.SMTP( MAIL_HOST, MAIL_PORT)
    except:
        mailServer = 0
        f_log.write (time.asctime(time.localtime()))
        f_log.write (" Errore in connessione al server SMTP:\n")


    #f_log.write (time.asctime(time.localtime()))
    #f_log.write (" HOST " + MAIL_HOST + " Porta: " + MAIL_PORT + "\n")
    if (mailServer):
        mailServer.ehlo()
        mailServer.starttls()
        #mailServer.ehlo()
        ##### ATTENZIONE - NON COMMENTARE LA RIGA SOTTO PER AUTENTICARSI SU SHAZAM O GMAIL 
        if AUTENTICAZIONE == "1":
            mailServer.login(User, Password)
        mailServer.sendmail(User, recipient, msg.as_string())
        mailServer.close()
        f_log.write (time.asctime(time.localtime()))
        f_log.write (" Email Inviata a " + recipient + "\n")
        #print('Sent email to %s' % recipient)
def getAttachment(attachmentFilePath):
    contentType, encoding = mimetypes.guess_type(attachmentFilePath)

    if contentType is None or encoding is not None:
        contentType = 'application/octet-stream'

    mainType, subType = contentType.split('/', 1)
    file = open(attachmentFilePath, 'rb')

    if mainType == 'text':
        attachment = MIMEText(file.read())
    elif mainType == 'message':
        attachment = email.message_from_file(file)
    elif mainType == 'image':
        attachment = MIMEImage(file.read(),_subType=subType)
    elif mainType == 'audio':
        attachment = MIMEAudio(file.read(),_subType=subType)
    else:
        attachment = MIMEBase(mainType, subType)
    
    attachment.set_payload(file.read())
    encode_base64(attachment)

    file.close()

    attachment.add_header('Content-Disposition', 'attachment',   filename=os.path.basename(attachmentFilePath))
    #f_log.write (time.asctime(time.localtime()))
    #f_log.write (" Allegato Email ok\n")
    return attachment


###########################################################################
    
NomeFile = str (lista[1])
#print(str(NomeFile) + "\n")
NomeFileNoExt = NomeFile.split(".")
#print(str (NomeFileNoExt[0]) + "\n")


from reportlab.pdfgen.canvas import Canvas
#c = Canvas('C:\\temp\\temp.pdf')

#check - routine nomefile
#NomePDF = str (NomeFileNoExt[0]) + ".pdf"


##################### Estrapolo il nome file del nuovo pdf ################
wrong_chars= ['\\','/',':','*','?','"','<','>','|']
contarighe = 0
for line in file(NomeFile,'r'):
    contarighe = contarighe + 1
    if contarighe == 1:
        uniLine = unicode(line, 'utf-8')
        #if uniLine.startswith ("ï»¿"):
        #    print("Problema nel file")
        #f_log.write (time.asctime(time.localtime()))
        #f_log.write (" Nome PDF: " + uniLine[0:26] + ".pdf\n")
        uniLine = uniLine [0:60]
        NomePDFNoPath = uniLine.rstrip() + ".pdf"
        for i in NomePDFNoPath:
            if i in wrong_chars:
                f_log.write (time.asctime(time.localtime()))
                f_log.write (" Impossibile utilizzare il determinato nome file pdf. Settato Default.\n")
                NomePDFNoPath = str(time.time()).replace('.','_') #"Default"
        NomePDF = DIR_RIFERIMENTO +"\\Out\\" + NomePDFNoPath #+ ".pdf"



c = Canvas(NomePDF)
c.setAuthor("Tarros S.P.A")
c.setTitle("Documento QTZ")
f_log.write (time.asctime(time.localtime()))
f_log.write (" Processo File: " + NomePDFNoPath + "\n" )
f_log.write (time.asctime(time.localtime()))
f_log.write (" Percorso File: " + NomePDF + "\n" )
    
    
import reportlab.rl_config
reportlab.rl_config.warnOnMissingFontGlyphs = 0
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


############ Registro vari fonts ####################################

#pdfmetrics.registerFont(TTFont('Vera', 'Vera.ttf'))
#pdfmetrics.registerFont(TTFont('VeraBd', 'VeraBd.ttf'))
#pdfmetrics.registerFont(TTFont('VeraIt', 'VeraIt.ttf'))
#pdfmetrics.registerFont(TTFont('VeraBI', 'VeraBI.ttf'))

pdfmetrics.registerFont(TTFont('LucidaConsole', 'lucon.ttf'))

#pdfmetrics.registerFont(TTFont('FreeMono', 'FreeMono.ttf'))

#pdfmetrics.registerFont(TTFont('DejaVuSansMono', 'DejaVuSansMono.ttf'))
#pdfmetrics.registerFont(TTFont('DejaVuSansMono-Bold', 'DejaVuSansMono-Bold.ttf'))
#pdfmetrics.registerFont(TTFont('DejaVuSansMono-BoldOblique', 'DejaVuSansMono-BoldOblique.ttf'))
#pdfmetrics.registerFont(TTFont('DejaVuSansMono-Oblique', 'DejaVuSansMono-Oblique.ttf'))
########################################################################

############ Apro e leggo il file ini per le parole chiave ############
try:
    f_keys_ini = open("TarrosKeys.ini",'r')
    #print ("File ini",f_ini.name, "aperto..")
except IOError:
    f_log.write (time.asctime(time.localtime()))
    f_log.write (" Impossibile aprire il file ini: TarrosKeys.ini\n")
    sys.exit (2)

parole_chiave = []
for parola in f_keys_ini:
    parola = parola.replace("\"","")
    parole_chiave.append(parola[:-1])


#print(str(parole_chiave) + "\n")
f_keys_ini.close() 


y = 810
#y = 750
counter = 1
formattazione = "off"
pagina = 1
giallo = PCMYKColor (0,34,94,0)
blue = PCMYKColor (100,91,30,22)

#### disegno la linea blu e logo sulla prima pagina
def disegnaSfondo(logo):
    c.setStrokeColor (blue)
    c.setLineWidth (1)
    c.line(30,743,565,743)
    c.setFillColor(blue)
    c.setFontSize(12)
    c.drawString(258,716,"Offer Code")
    try:
        c.drawImage(DIR_RIFERIMENTO +"\\grf\\" + logo ,232,750,123,85,preserveAspectRatio=True)
        f_log.write (time.asctime(time.localtime()))
        f_log.write (" Inserito logo: " + DIR_RIFERIMENTO +"\\grf\\" + logo + "\n")
    except:
        c.drawImage(DIR_RIFERIMENTO +"\\grf\\LOGO_TARROS_COLORE2.jpg",232,750,123,85,preserveAspectRatio=True)
        f_log.write (time.asctime(time.localtime()))
        f_log.write (" Errore inserendo il logo: " + DIR_RIFERIMENTO +"\\grf\\" + logo + " - Inserito il logo default\n")
    c.setStrokeColor(giallo)
    c.setLineWidth(1)
    c.rect (244,698,102,30,stroke = 1) 




for line in file(NomeFile,'r'):
    uniLine = unicode(line, 'utf-8')
    formattazione = "off"
    
    if counter == 73:
        pagina = pagina + 1
        c.setFont('LucidaConsole', 8)
        c.showPage()
        counter = 1
        y = 810 
        #y = 750 
 
    
    for parola in parole_chiave:
        if uniLine.startswith(parola):
            formattazione = "on"
    
    # nascondo la prima riga e riposiziono il numero pagina
    if counter == 1:
        Pag = uniLine[101:109]
        uniLine = " "
        c.setFont('LucidaConsole', 8)
        c.setFillColorRGB (0,0,0)
        c.drawString (504,26,Pag)
        #disegnaSfondo()
        #print("counter:" + str(counter) + " pag: " + str(Pag)+"\n")

    if counter == 4:
        logo = uniLine[0:]
        logo = logo.rstrip ()
        uniLine = " "
        #c.setFont('LucidaConsole', 8)
        #c.setFillColorRGB (0,0,0)
        #c.drawString (504,26,Pag)
        disegnaSfondo(logo)
        #print("counter:" + str(counter) + " pag: " + str(Pag)+"\n")    
    
    if counter == 12:
        uniLine = uniLine[4:]
    
    #if (pagina == 1) and (counter == 2): #ho eliminato il check su pagina perchè si ripetono su ogni pagina..
    if counter == 2:
        indirizzoEmail = uniLine[1:]
        indirizzoEmail = indirizzoEmail.rstrip()
        if len (indirizzoEmail) == 0:
            POSTA = "NO"
            #indirizzoEmail = MAIL_TO
        else:
            POSTA = "SI"
        f_log.write (time.asctime(time.localtime()))
        f_log.write (" Indirizzo Email su riga 2: " + indirizzoEmail + "\n")
        uniLine = " "
     

    #if (pagina == 1) and (counter == 3): #ho eliminato il check su pagina perchè si ripetono su ogni pagina..
    if counter == 3:
        Percorso = uniLine[1:]
        Percorso = Percorso.rstrip()
        if len (Percorso) == 0:
            Percorso = DIR_RIFERIMENTO
            f_log.write (time.asctime(time.localtime()))
            f_log.write (" Percorso e DIR_RIFERIMENTO sono uguali\n")
        else:
            f_log.write (time.asctime(time.localtime()))
            f_log.write (" Percorso e DIR_RIFERIMENTO sono DIVERSI\n")
            f_log.write (time.asctime(time.localtime()))
            f_log.write (" Percorso : " + Percorso + "\n")
        uniLine = " "
    #if (pagina == 1) and (counter == 5): #ho eliminato il check su pagina perchè si ripetono su ogni pagina..
    if counter == 5:
        SoggettoEmail = uniLine[0:]
        SoggettoEmail = SoggettoEmail.rstrip()
        if len (SoggettoEmail) == 0:
            SoggettoEmail = "Offer Code: "
            f_log.write (time.asctime(time.localtime()))
            f_log.write (" Soggetto email di default\n")
        else:
            f_log.write (time.asctime(time.localtime()))
            f_log.write (" Soggetto email estratto dai dati\n")
        uniLine = " "
 
    if uniLine.startswith ("        VIASPAZIVIASPAZI"):
        uniLine = " "


    if uniLine.startswith("Best regards") or uniLine.startswith (" Best regards"):
        ############
        immagine = uniLine[56:]
        immagine = immagine.rstrip()
        try:
            #c.drawImage (immagine,40,y-40,200,50,preserveAspectRatio=True)
            c.drawImage (immagine,40,y-55,241,50,preserveAspectRatio=True)
        except:
            f_log.write (time.asctime(time.localtime()))
            f_log.write(" Errore inserendo immagine dinamica " + str(immagine) + "\n")
            #print ("Errore inserendo immagine dinamica " + str(immagine) + "\n")
        #c.drawImage ("zan.jpg",40,y-40,200,50,preserveAspectRatio=True)
        uniLine = uniLine[0:40] 
        print ("immagine inserita a y = " + str (y))

    #prefisso = str (y)
    #uniLine = prefisso.rjust (3,"0") + uniLine
    #counter = counter + 1

   
    #print ("Scritte " + str (counter)+ " righe.. su pagina " + str (pagina))
    #f_log.write ("Scritte " + str (counter)+ " righe.. su pagina " + str (pagina) + "\n")
    
    if formattazione == "on":
        c.setFillColor(giallo)
        c.setStrokeColor(blue)
        c.setLineWidth(0.5)
        c.rect (30,y - 2,535, 9, stroke = 1, fill = 1) 
        c.setFont ('LucidaConsole', 8)
        c.setFillColor(blue)
        #c.drawString(40, y, uniLine.strip())
        c.drawString(40, y, uniLine[:-1])
    else:
        if (pagina == 1) and ((counter == 16) or (counter == 17)):
            c.setFont ('LucidaConsole',10)
            c.setFillColorRGB (0,0,0)
            #c.drawString(40, y, uniLine.strip())
            c.drawString(30, y, uniLine[:-1])
        else:
            c.setFont('LucidaConsole', 8)
            c.setFillColorRGB (0,0,0)
            #c.drawString(40, y, uniLine.strip())
            c.drawString(40, y, uniLine[:-1])

    counter = counter + 1
    y = y - 10


'''
if y < 40:
    c.drawImage ("zan.jpg",40,y,200,50,preserveAspectRatio=True)
else:
    print ("y = " + str (y))
    print ("Possibile problema di spazio a posizionare l'immagine?")
try:
    c.drawImage(DIR_RIFERIMENTO + "\\Grf\\LOGO_TARROS_COLORE2.JPG", 700,100,200,80,preserveAspectRatio=True) 
except e:
    print (e)
'''


c.showPage()
c.save()

if pagina > 1:
    #print ("TOTALE Scritte " +str (pagina) + " pagine")
    f_log.write (time.asctime(time.localtime()))
    f_log.write(" TOTALE Scritte " +str (pagina) + " pagine\n")
else:
    #print ("TOTALE Scritta " +str (pagina) + " pagina")
    f_log.write (time.asctime(time.localtime()))
    f_log.write(" TOTALE Scritta " +str (pagina) + " pagina\n")

#print ("pdf scritto")
f_log.write (time.asctime(time.localtime()))
f_log.write (" Pdf Scritto\n")


#check
#os.remove (NomeFile)

############ INVIO EMAIL ############
corpo_email = """Attached Offer Code.
Best regards.
"""

if POSTA == "SI":
    sendMail (SoggettoEmail +" - File " + str (NomePDFNoPath) + "",corpo_email,NomePDF)
else:
    f_log.write (time.asctime(time.localtime()))
    f_log.write (" Nessun Indirizzo di Posta nei dati\n")

if Percorso != DIR_RIFERIMENTO :
    if os.path.exists(Percorso):
        shutil.copy(NomePDF,Percorso)
        f_log.write (time.asctime(time.localtime()))
        f_log.write(" Pdf copiato su dir: " + Percorso + "\n")
        os.remove(NomePDF)
    else:
        f_log.write (time.asctime(time.localtime()))
        f_log.write(" Problemi nel copiare i file su: " + Percorso + "\n")
else:
    f_log.write (time.asctime(time.localtime()))
    f_log.write(" Percorso non presente nei dati\n")
 
f_log.write (time.asctime(time.localtime()))
f_log.write (" Fine Elaborazione\n")
f_log.close()
