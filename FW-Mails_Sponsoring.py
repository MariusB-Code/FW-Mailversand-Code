from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import smtplib
import openpyxl
import pandas as pd
from docx import Document
import os

# Daten für Server, Email-Login
server_name = "server-Name"   #hostname of SMTP-server
portNr = 587                            #server port for SMTP-connection (usually 25 or 587)
sender = "sender-Email-address"     #username for server-login (usually same as sender-email address)
password = "server-password"               #password for server-login
# Pfadangaben für Anhänge
xlsx_path = "./Mailingliste_Test.xlsx"          #Pfad zur xlsx-Datei mit E-Mail-Adressen, Anrede
docx_path = "./FW26_Mailvorlage_Anfrage.docx"    #Pfad für Plain-text-version der email
html_path = "./FW26_Mailvorlage_Anfrage.htm"     #Pfad der HTML-Version der email !Muss UTF-8 sein!
attach1_path = "./Sponsoring_Dossier_Frackwoche26.pdf"      #Pfad zum Anhang 1
attach2_path = "./Sponsoring_Formular_Frackwoche26.pdf"     #Pfad zum Anhang 2
image1_path = "./FW26_Mailvorlage_Anfrage-Dateien/image001.jpg"    #Pfad zum Bild 1 für html
image2_path = "./FW26_Mailvorlage_Anfrage-Dateien/image002.jpg"    #Pfad zum Bild 2 für html
image3_path = "./FW26_Mailvorlage_Anfrage-Dateien/image003.jpg"    #Pfad zum Bild 3 für html
image4_path = "./FW26_Mailvorlage_Anfrage-Dateien/image004.jpg"    #Pfad zum Bild 4 für html
#image5_path = "./FW26_Mailvorlage_Anfrage-Dateien/image005.jpg"    #Pfad zum Bild 5 für html
 # Betreff der Nachricht
subject = "Betreff"
# sicherheitssperre: bei "False" werden E-Mails simuliert
emailsSendable = False

#make attachements and images into array
attachArr = []
attachArr.append(attach1_path)
attachArr.append(attach2_path)
imgArr = []
imgArr.append(image1_path)
imgArr.append(image2_path)
imgArr.append(image3_path)
imgArr.append(image4_path)
#imgArr.append(image5_path)

# other Variables
numRows = 0
numSavedMail=0

# Function createEmail: assembles the Multipart email with:
# Sender, reciepent, cc, subject
# "mixed" encoding for attaments
#   "related" encoding for HTML with embedded images
#       "alternative" encoding for plain-text and html versions of the email
#
#MIMEMultipart structure:
#multipart/mixed                ← Root (because of attachments)
#├── multipart/related          ← For HTML + embedded images
#│   ├── multipart/alternative  ← For text and HTML versions
#│   │   ├── text/plain         ← Plain text body
#│   │   └── text/html          ← HTML body
#│   └── image/jpeg             ← Embedded image(s)
#├── application/pdf            ← Attachment(s)
#└── other attachments...
#--------------------------------------------------------------
def createEmail(sender,recipient,cc,subject,text,html,attachments,images):

    # Create a multipart message with tag "mixed", to support attachements
    msg = MIMEMultipart("mixed")
    msg["From"] = sender
    msg["To"] = recipient
    if isinstance(cc,str): #check if cc is sting (if no cc then python defaults to float)
        msg["Cc"] = cc
    msg["Subject"] = subject


    # create MIMEMultipart subfields with tags "related" for html with embedded images
    #   and "alternative" for plain text/html versions of the email
    rel = MIMEMultipart("related")
    alt = MIMEMultipart("alternative")

    # attach plain-text and HTML versons of the email ("alternative")
    # add HTML/plain-text parts to MIMEMultipart message with utf-8 encoding for äöü
    #   the recieving email client will try to render the last part first
    alt.attach(MIMEText(text,"plain","utf-8"))
    alt.attach(MIMEText(html, "html","utf-8"))
    rel.attach(alt)

    #add embedded images to the html-part ("related")
    if images:
        i = 1
        for image in images:
            with open(image,"rb") as image:
                imgPart = MIMEImage(image.read())
                cid = "image00"+str(i)
                imgPart.add_header('Content-ID',cid)
                rel.attach(imgPart)
                i+=1
                #print(f"Image added: {image}")
    msg.attach(rel)

    #add attachements ("mixed")
    if attachments:
        for attachment in attachments:
            with open(attachment, 'rb') as f:
                file = MIMEApplication(f.read(), name=os.path.basename(attachment))
                file['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment)}"'
                msg.attach(file)

    return msg


# Function sendEmail: sends the email, prints confirmation
#--------------------------------------------------------------
def sendEmail(msg,sender,password,server_name,portNr,recipient,cc,numMail):
    if emailsSendable:
        server = smtplib.SMTP(server_name, portNr)
        server.starttls()
        server.login(sender,password)
        server.send_message(msg,sender,recipient)
        server.quit
        print(f"message {numMail} of {numRows} sent to: {recipient}, cc:{cc}")
    else:
        print(f"simulated message {numMail} of {numRows} sent to: {recipient}, cc:{cc}")


#Function: checkDocPaths: check if the Paths to the .docx, .htm and .xlsx-files are correct
#--------------------------------------------------------------
def checkDocPaths():
    if not os.access(docx_path,os.F_OK):
        print(f"\033[1;37;41m Dateipfad Falsch! Prüfen, ob Dateipfad von .docx-Datei richtig ist\033[0;0m")
        return False
    if not os.access(xlsx_path,os.F_OK):
        print(f"\033[1;37;41m Dateipfad Falsch! Prüfen, ob Dateipfad von .xlsx-Datei richtig ist\033[0;0m")
        return False
    if not os.access(html_path,os.F_OK):
        print(f"\033[1;37;41m Dateipfad Falsch! Prüfen, ob Dateipfad von HTML-Datei richtig ist\033[0;0m")
        return False
    print(f"\033[0;32;40m Überprüfung der Dateipfade erfolgreich\033[0;0m")
    return True


#Function: checkFilePaths: check if the Attachement file paths are correct
#--------------------------------------------------------------
def checkFilePaths(filePathArr):
    i=0
    for filePath in filePathArr:
        if not os.access(filePath,os.F_OK):
            print(f"\033[1;37;41m Dateipfad Falsch! Prüfen, ob Dateipfad von Anhang {i+1} richtig ist\033[0;0m")
            return False
    i+=1
    print(f"\033[0;32;40m Überprüfung der Anhangdateipfade erfolgreich\033[0;0m")
    return True


#Function: checkFilePaths: check if the Image file paths for embdeded images are correct
#--------------------------------------------------------------
def checkImagePaths(imagePathArr):
    i=0
    for imgPath in imagePathArr:
        if not os.access(imgPath,os.F_OK):
            print(f"\033[1;37;41m Dateipfad Falsch! Prüfen, ob Dateipfad von Bild {i+1} richtig ist\033[0;0m")
            return False
    i+=1
    print(f"\033[0;32;40m Überprüfung der Bildpfade erfolgreich\033[0;0m")
    return True


#Function: checkFields: check if all  nessecary fields are correctly filled in, return True if so
#--------------------------------------------------------------
def checkFields(df):
    for i,row in df.iterrows():
        #check whether Instances "Mail" and "Anrede" contain something
        if not isinstance(row["Mail"],str):
            print(f"\033[1;37;41m Mail ist kein Sting! Prüfen, ob 'Mail' in Zeile {i+2} Inhalt hat\033[0;0m")
            return False
        if not isinstance(row["Anrede"],str):
            print(f"\033[1;37;41m Anrede ist kein Sting! Prüfen, ob 'Anrede' in Zeile {i+2} Inhalt hat\033[0;0m")
            return False
        #check whether mail-addresses contains ASCII-Letters only
        if not (row["Mail"].isascii()):
            print(f"\033[1;37;41m Mailadresse in Zeile {i+2} enthält nicht-ASCII-Zeichen! Prüfen, ob alles ASCII-Charaktere sind! \033[0;0m")
            print(f"\033[0;31;40m Email-Adresse nochmals abtippen, danach sollte das Problem behoben sein\033[0;0m")
            return False
        #check whether mail-addresses contains space-character
        if " " in row["Mail"]:
            print(f"\033[1;37;41m Mailadresse in Zeile {i+2} enthält leerzeichen! Leerzeichen entfernen\033[0;0m")
            return False
        #check whether mail-addresses contains @-character
        if not "@" in row["Mail"]:
            print(f"\033[1;37;41m Mailadresse in Zeile {i+2} enthält kein @-Zeichen! Mailadresse prüfen\033[0;0m")
            return False
        #check wheter cc-adresses (if present) conatains ASCII only, no space-cahr and @-char
        if isinstance(row["Cc"],str): #check if cc is sting (if no cc then python defaults to float)
            if not (row["Cc"].isascii()):
                print(f"\033[1;37;41m CC-Adresse in Zeile {i+2} enthält nicht-ASCII-Zeichen! Prüfen, ob alles ASCII-Charaktere sind! \033[0;0m")
                print(f"\033[0;31;40m CC-Adresse nochmals abtippen, danach sollte das Problem behoben sein\033[0;0m")
                return False
            if " " in row["Cc"]:
                print(f"\033[1;37;41m CC-Adresse in Zeile {i+2} enthält leerzeichen! Leerzeichen entfernen\033[0;0m")
                return False
            if not "@" in row["Cc"]:
                print(f"\033[1;37;41m CC-Adresse in Zeile {i+2} enthält kein @-Zeichen! CC-Adresse prüfen\033[0;0m")
                return False

    print(f"\033[0;32;40m Überprüfung der 'Anrede', 'Cc' und 'Mail'-Felder erfolgreich\033[0;0m")
    return True


#main function: read csv, docx and htm files, make email
#--------------------------------------------------------------

# check if .docx, .xlsx and .html-file exist
docPathsCorrect = checkDocPaths()
attachPathCorrect = checkFilePaths(attachArr)
imagePathCorrect = checkImagePaths(imgArr)

# read excel file, check if nessecary fields filled in, get number of rows
if docPathsCorrect:
    df = pd.read_excel(xlsx_path)
    #print(df)
    fieldsCorrect = checkFields(df)
    numRows = df.shape[0]

#start creating and sending Mails
if (docPathsCorrect and attachPathCorrect and imagePathCorrect and fieldsCorrect):

    numMail = 0
    numSavedMail=0
    #check if a save-file exists
    if os.path.isfile("./MassenMails_saveRow.txt"):
        with open("./MassenMails_saveRow.txt") as f:
            numSavedMail = int(f.read())

        print(f"begonnener Prozess erkannt: Startet von Mail Nr.{numSavedMail}")

    for index, row in df.iterrows():
        numMail +=1
        #if numMail is larger than saved mail, continue, else cycle with no/op
        if (numMail > numSavedMail):
        
            anrede = row["Anrede"]
            recipient = row["Mail"]
            cc = row["Cc"]
            #print(f"anrede: {anrede}, enpfänger: {recipient}, Cc: {cc} (nan falls nicht vorhanden)")

            # create plain text version of email:
            #   open .docx file, replace [Anrede] for plain text-version of email
            text = []
            doc = Document(docx_path)
            for paragraph in doc.paragraphs:
                #print(paragraph.text)
                if "[Anrede]" in paragraph.text:
                    paragraph.text = anrede
                text.append(paragraph.text)

            text = '\n'.join(text)
            #print(text)

            # create html version of email:
            #   open .htm file, replace [Anrede] for html-version of email
            html = []
            with open(html_path,'r',encoding="utf-8") as f:
                html = f.read()
                html = html.replace("[Anrede]",anrede)
            #print(html)

            msg = createEmail(sender,recipient,cc,subject,text,html,attachArr,imgArr)
            try:
                sendEmail(msg,sender,password,server_name,portNr,recipient,cc,numMail)
            except:
                print(f"\033[0;31;40m Email-wegen Fehler abgebrochen. Fehler beheben und neu starten\033[0;0m")
                print(f"\033[0;31;40m Falls Fehler wegen 'Too many E-mails' ca. 1h warten und erneut versuchen\033[0;0m")
                f=open("./MassenMails_saveRow.txt", "w")
                f.write(str(numMail))
                print(f"\033[0;33;40m Anzahl versandter Mails gespeichert: startet nächstes mal von Mail Nr. {numMail}\033[0;0m")
        
        (f"\033[0;32;40m Alle Mails erfolgreich versendet.\033[0;0m")
        if os.path.isfile("./MassenMails_saveRow.txt"):
            os.remove("./MassenMails_saveRow.txt")
            print("save-Datei entfernt")

else:
    print(f"\033[0;31;40m Email-Übertragung wurde nicht gestartet\033[0;0m")